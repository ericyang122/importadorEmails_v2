import argparse
import re
import time
import json
import os
import sys
import unicodedata
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from decimal import Decimal, InvalidOperation
import pandas as pd
import requests
from dotenv import load_dotenv

# Garante UTF-8 na saida do script. Sem isso, rodar o confio.py direto no
# terminal do Windows (que usa cp1252) trava com UnicodeEncodeError ao imprimir
# simbolos como ✓/✗. Pela interface ja vem UTF-8; aqui blindamos os dois casos.
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

def parse_args():
    parser = argparse.ArgumentParser(description="Importa leads de uma planilha para o Sigavi.")
    parser.add_argument(
        "--excel",
        default=os.getenv("SIGAVI_EXCEL", "./abertos/abertos_cora.xlsx"),
        help="Caminho da planilha Excel a importar.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Executa o navegador sem janela visivel.",
    )
    parser.add_argument(
        "--result-dir",
        default=os.getenv("SIGAVI_RESULT_DIR", "resultados"),
        help="Pasta onde a planilha de resultado sera salva.",
    )
    parser.add_argument(
        "--mode",
        choices=("consulta", "cadastro"),
        default=os.getenv("SIGAVI_MODE", "consulta"),
        help="Modo de execucao: consulta busca telefones por email; cadastro cadastra leads com telefone.",
    )
    parser.add_argument(
        "--stop-file",
        default=os.getenv("SIGAVI_STOP_FILE"),
        help="Arquivo sentinela usado pela interface para pedir parada segura.",
    )
    parser.add_argument(
        "--progress-file",
        default=os.getenv("SIGAVI_PROGRESS_FILE"),
        help="Arquivo JSON de progresso da execucao.",
    )
    parser.add_argument(
        "--autosave-every",
        type=int,
        default=int(os.getenv("SIGAVI_AUTOSAVE_EVERY", "10")),
        help="Quantidade de linhas processadas entre salvamentos dos Excels parciais.",
    )
    return parser.parse_args()


ARGS = parse_args()
load_dotenv()
SIGAVI_LOGIN = os.getenv("SIGAVI_LOGIN")
SIGAVI_SENHA = os.getenv("SIGAVI_SENHA")
HEADLESS = ARGS.headless or os.getenv("SIGAVI_HEADLESS", "").strip().lower() in {"1", "true", "yes", "sim"}
RESULT_DIR = ARGS.result_dir
MODE = ARGS.mode
STOP_FILE = ARGS.stop_file
AUTOSAVE_EVERY = max(1, ARGS.autosave_every)

if not SIGAVI_LOGIN or not SIGAVI_SENHA:
    print("ERRO: informe SIGAVI_LOGIN e SIGAVI_SENHA no .env ou no ambiente da execucao.")
    raise SystemExit(1)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementNotInteractableException,
    ElementClickInterceptedException,
    NoSuchElementException,
    InvalidSessionIdException,
    WebDriverException,
)

BASE_URL = "https://abyara.sigavi360.com.br"
# Quantas buscas por email rodam em paralelo no modo consulta.
# 8 e o padrao; o retry anti-throttle recupera respostas vazias, entao nao
# perde telefone mesmo se o Sigavi limitar. Baixe via .env se ficar lento.
CONSULTA_WORKERS = max(1, int(os.getenv("SIGAVI_CONSULTA_WORKERS", "8")))
# Quantas vezes retentar uma busca com resposta suspeita/erro antes de desistir.
CONSULTA_TENTATIVAS = max(1, int(os.getenv("SIGAVI_CONSULTA_TENTATIVAS", "3")))
# Base (segundos) do backoff entre tentativas; cresce a cada tentativa.
CONSULTA_BACKOFF = float(os.getenv("SIGAVI_CONSULTA_BACKOFF", "1.5"))
# Fator de velocidade do cadastro: multiplica todas as pausas de UI do Selenium.
# 1.0 = comportamento original. Abaixe (ex: 0.6) para acelerar; suba se a tela
# do Sigavi estiver lenta e o cadastro comecar a falhar.
CADASTRO_DELAY = max(0.1, float(os.getenv("SIGAVI_CADASTRO_DELAY", "1.0")))
# Termo digitado no modal "Imovel de origem" para escolher o empreendimento do
# cadastro. Trocavel pelo .env (SIGAVI_EMPREENDIMENTO=arvo) sem mexer no codigo.
EMPREENDIMENTO_BUSCA = (os.getenv("SIGAVI_EMPREENDIMENTO", "arvo").strip() or "arvo")
# Protege a renovacao de cookies/token (o driver Selenium nao e thread-safe).
_token_lock = threading.Lock()


def pausa_cadastro(segundos):
    """Pausa proporcional ao CADASTRO_DELAY (permite calibrar a velocidade)."""
    time.sleep(segundos * CADASTRO_DELAY)

def carregar_progresso():
    if os.path.exists(PROGRESSO_FILE):
        with open(PROGRESSO_FILE, 'r', encoding='utf-8') as f:
            dados = json.load(f)
        print(f"Retomando do índice {dados['ultimo_index'] + 1} (progresso salvo encontrado).")
        return (
            dados.get('ultimo_index', -1),
            dados.get('resultados_email', []),
            dados.get('resultados_cadastro', []),
        )
    return -1, [], []

def salvar_progresso(ultimo_index, resultados_email, resultados_cadastro):
    os.makedirs(os.path.dirname(PROGRESSO_FILE) or '.', exist_ok=True)
    with open(PROGRESSO_FILE, 'w', encoding='utf-8') as f:
        json.dump({
            'ultimo_index': ultimo_index,
            'modo': MODE,
            'resultados_email': resultados_email,
            'resultados_cadastro': resultados_cadastro,
        }, f, ensure_ascii=False)

# =========================
# DICIONÁRIO CORRETORES (carregado do corretores.json)
# =========================
corretores_gerentes = {}
try:
    with open('corretores.json', 'r', encoding='utf-8') as f:
        corretores_json = json.load(f)
        if isinstance(corretores_json, list) and corretores_json:
            corretores_gerentes = corretores_json[0].get('corretoresEquipes', {})
        elif isinstance(corretores_json, dict):
            corretores_gerentes = corretores_json.get('corretoresEquipes', {})
    print(f"Carregados {len(corretores_gerentes)} corretores do corretores.json")
except Exception as e:
    print(f"ERRO: Não foi possível carregar corretores.json: {e}")
    raise SystemExit(1)

# =========================
# LEITURA DA PLANILHA
# =========================
arquivo_excel = ARGS.excel
if not os.path.exists(arquivo_excel):
    print(f"ERRO: planilha nao encontrada: {arquivo_excel}")
    raise SystemExit(1)

excel_file     = pd.ExcelFile(arquivo_excel)
nome_planilha  = excel_file.sheet_names[0]
df             = pd.read_excel(arquivo_excel, sheet_name=nome_planilha)
print(f"Arquivo carregado: {arquivo_excel} | Aba: '{nome_planilha}' | {len(df)} linhas")

df = df.rename(columns={
    # formato antigo
    'CORRETOR ORIGEM': 'CORRETOR DE ORIGEM',
    'TELEFONE'       : 'FONE2',
    'NOME COMPLETO'  : 'NOME',
    # formato Base Cora
    'nome_cliente'   : 'NOME',
    'celular'        : 'FONE2',
    'corretor'       : 'CORRETOR DE ORIGEM',
    'origem'         : 'TIPO PLANTAO',
    'gerente'        : 'GERENTE',
})

# celular vem como float (ex: 1.199622e+10) → converte para string de dígitos
def normalizar_telefone_planilha(valor):
    if valor is None or pd.isna(valor):
        return ''
    texto = str(valor).strip()
    if texto.lower() in ('', 'nan', 'none'):
        return ''
    try:
        numero = Decimal(texto)
        if numero.is_finite() and numero == numero.to_integral_value():
            return str(int(numero))
    except InvalidOperation:
        pass
    return re.sub(r'\D', '', texto)


if 'FONE2' in df.columns:
    df['FONE2'] = df['FONE2'].apply(normalizar_telefone_planilha)

_nome_excel    = os.path.splitext(os.path.basename(arquivo_excel.lstrip('./')))[0]
PROGRESSO_FILE = ARGS.progress_file or os.path.join(RESULT_DIR, f'progresso_{_nome_excel}_{MODE}.json')

def parada_solicitada():
    return bool(STOP_FILE and os.path.exists(STOP_FILE))


def _resultado_file(sufixo):
    return os.path.join(RESULT_DIR, f'resultado_{_nome_excel}_{sufixo}.xlsx')

# =========================
# CHROME + HELPERS
# =========================
def criar_driver():
    # Navegador configuravel: SIGAVI_BROWSER=edge (default, Windows) ou chrome (Linux/VM).
    browser = os.getenv("SIGAVI_BROWSER", "edge").strip().lower()
    usa_chrome = browser in {"chrome", "chromium"}

    options = webdriver.ChromeOptions() if usa_chrome else webdriver.EdgeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    if HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    d = webdriver.Chrome(options=options) if usa_chrome else webdriver.Edge(options=options)
    if HEADLESS:
        d.set_window_size(1920, 1080)
    else:
        d.maximize_window()
    return d

driver = criar_driver()
wait = WebDriverWait(driver, 20)

def reconectar_browser():
    global driver, wait
    print("Browser caiu. Reconectando...")
    try:
        driver.quit()
    except Exception:
        pass
    time.sleep(3)
    driver = criar_driver()
    wait = WebDriverWait(driver, 20)
    time.sleep(2)
    # Re-login
    driver.get("https://abyara.sigavi360.com.br/Acesso/Login?ReturnUrl=%2F")
    time.sleep(4)
    WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[1]/div/input"))   
    ).send_keys(SIGAVI_LOGIN)
    WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[2]/div/input"))
    ).send_keys(SIGAVI_SENHA)
    WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[3]/div/button"))
    ).click()
    time.sleep(3)
    print("Reconectado e logado com sucesso.")

def wait_visible(locator, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located(locator)
    )

def wait_clickable(locator, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(locator)
    )

def scroll_into_view(el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)

def safe_click(locator):
    el = wait_visible(locator)
    scroll_into_view(el)
    try:
        wait_clickable(locator)
        el.click()
    except (ElementClickInterceptedException, ElementNotInteractableException, TimeoutException):
        driver.execute_script("arguments[0].click();", el)

def sessao_caiu_para_login():
    """True se o Sigavi redirecionou pra tela de login (sessao expirada ou derrubada)."""
    try:
        url = (driver.current_url or '').lower()
    except WebDriverException:
        return False
    return 'login' in url or '/acesso' in url

def relogar_sigavi():
    # O Sigavi derruba a sessao no meio do job (timeout de inatividade ou
    # sessao concorrente em outro lugar) e passa a redirecionar /CRM/Fac pro
    # login. Sem relogar aqui, TODO lead a partir da queda falhava com
    # "Pagina /CRM/Fac nao carregou" ate o fim da planilha (job de 2026-06-11).
    print("Sessao do Sigavi caiu (redirecionado pro login). Relogando...")
    driver.get("https://abyara.sigavi360.com.br/Acesso/Login?ReturnUrl=%2F")
    time.sleep(3)
    wait_visible((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[1]/div/input"), timeout=30)\
        .send_keys(SIGAVI_LOGIN)
    wait_visible((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[2]/div/input"), timeout=30)\
        .send_keys(SIGAVI_SENHA)
    safe_click((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[3]/div/button"))
    time.sleep(3)
    driver.get(BASE_URL + "/")
    time.sleep(2)
    if sessao_caiu_para_login():
        raise RuntimeError(
            "Relogin no Sigavi nao foi aceito (voltou pra tela de login). "
            "Possivel sessao concorrente aberta em outro lugar ou captcha."
        )
    print("Relogin no Sigavi concluido.")

# =========================
# HELPERS ROBUSTOS DE COMBO/INPUT (portados da automacao do Pedro, que valida
# cada selecao em vez de mandar send_keys as cegas e torcer pra ter pegado)
# =========================
def texto_elemento(el):
    return (driver.execute_script(
        "return arguments[0].value || arguments[0].innerText || arguments[0].textContent || '';",
        el,
    ) or '').strip()

def combo_preenchido(locator, valor_esperado=''):
    """Confere se o combo mostra o valor esperado (e nao '---' ou vazio)."""
    try:
        texto = texto_elemento(wait_visible(locator, timeout=5))
    except TimeoutException:
        return False
    texto_norm = normalizar_texto(texto)
    valor_norm = normalizar_texto(valor_esperado)
    if not texto_norm or texto.strip() == '---':
        return False
    return not valor_norm or valor_norm in texto_norm or texto_norm in valor_norm

def elementos_visiveis(locators):
    elementos = []
    for locator in locators:
        for el in driver.find_elements(*locator):
            try:
                if el.is_displayed() and el.is_enabled():
                    elementos.append(el)
            except Exception:
                pass
    return elementos

def dropdowns_combo_visiveis():
    return driver.execute_script("""
        return Array.from(document.querySelectorAll('div, ul, ol')).filter(function(el) {
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            if (style.display === 'none' || style.visibility === 'hidden') return false;
            if (rect.width < 100 || rect.width > 900 || rect.height < 35 || rect.height > 600) return false;
            if (el.scrollHeight <= el.clientHeight + 8) return false;
            const cls = String(el.className || '').toLowerCase();
            const txt = String(el.innerText || '').trim();
            const looksLikeCombo = cls.includes('select2') || cls.includes('dropdown') || cls.includes('result') || cls.includes('combo') || txt.split('\\n').length >= 3;
            return looksLikeCombo;
        });
    """)

def input_combo_aberto(timeout=8):
    locators = [
        (By.XPATH, "//div[contains(@class, 'select2-drop') and not(contains(@style, 'display: none'))]//input"),
        (By.XPATH, "//span[contains(@class, 'select2-container--open')]//input"),
        (By.XPATH, "//input[contains(@class, 'select2-input') or contains(@class, 'select2-search__field')]"),
    ]
    fim = time.time() + timeout
    while time.time() < fim:
        elementos = elementos_visiveis(locators)
        if elementos:
            return elementos[-1]
        time.sleep(0.2)
    raise TimeoutException("Input aberto do combo nao encontrado")

def clicar_opcao_combo(valor, timeout=8):
    valor_norm = normalizar_texto(valor)
    locators = [
        (By.XPATH, "//*[contains(@class, 'select2-result-label')]"),
        (By.XPATH, "//*[contains(@class, 'select2-results__option')]"),
        (By.XPATH, "//*[contains(@class, 'select2-result') and not(contains(@class, 'select2-searching'))]"),
        (By.XPATH, "//*[@role='option']"),
    ]
    fim = time.time() + timeout
    for dropdown in dropdowns_combo_visiveis():
        try:
            driver.execute_script("arguments[0].scrollTop = 0;", dropdown)
        except Exception:
            pass
    opcoes_vistas = set()
    while time.time() < fim:
        opcoes = []
        for el in elementos_visiveis(locators):
            texto = texto_elemento(el)
            texto_norm = normalizar_texto(texto)
            if not texto_norm:
                continue
            if len(texto_norm) <= 40:
                opcoes_vistas.add(texto_norm)
            if texto_norm == valor_norm:
                opcoes.insert(0, el)
            elif valor_norm in texto_norm or texto_norm in valor_norm:
                opcoes.append(el)
        if opcoes:
            opcao = opcoes[0]
            scroll_into_view(opcao)
            try:
                opcao.click()
            except Exception:
                driver.execute_script("arguments[0].click();", opcao)
            return True
        for dropdown in dropdowns_combo_visiveis():
            try:
                candidatos = dropdown.find_elements(By.XPATH, ".//*[self::li or self::div or self::span or self::a][normalize-space()]")
                for el in candidatos:
                    try:
                        if not el.is_displayed():
                            continue
                        texto = texto_elemento(el)
                        texto_norm = normalizar_texto(texto)
                        if texto_norm and len(texto_norm) <= 40:
                            opcoes_vistas.add(texto_norm)
                        if texto_norm == valor_norm or valor_norm in texto_norm or texto_norm in valor_norm:
                            scroll_into_view(el)
                            try:
                                el.click()
                            except Exception:
                                driver.execute_script("arguments[0].click();", el)
                            return True
                    except Exception:
                        pass
                driver.execute_script("arguments[0].scrollTop = arguments[0].scrollTop + Math.max(80, arguments[0].clientHeight - 20);", dropdown)
            except Exception:
                pass
        time.sleep(0.4)
    if opcoes_vistas:
        print(f"Opcoes vistas ao procurar '{valor}': {', '.join(sorted(opcoes_vistas)[:12])}")
    return False

def selecionar_combo_texto(locator, valor, nome_campo, tentativas=4):
    """Seleciona uma opcao de combo digitando o texto e CONFERINDO se pegou."""
    for tentativa in range(1, tentativas + 1):
        safe_click(locator)
        time.sleep(0.6)
        try:
            try:
                campo_busca = input_combo_aberto(timeout=1.5)
                campo_busca.click()
                ActionChains(driver)\
                    .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)\
                    .send_keys(Keys.DELETE)\
                    .send_keys(valor)\
                    .perform()
                time.sleep(1.2)
            except TimeoutException:
                # Combo sem input de busca (estilo Kendo): digitar com o combo
                # focado pula direto pro item correspondente (era assim que o
                # codigo antigo selecionava CESARRICARDO digitando so CESAR).
                # A caca visual sozinha falha com opcoes fora da tela (ex.:
                # ELAINE, que exige rolar a lista ate o E).
                print(f"{nome_campo} sem input de busca; digitando '{valor}' direto no combo.")
                ActionChains(driver).send_keys(valor).perform()
                time.sleep(0.8)
                ActionChains(driver).send_keys(Keys.ENTER).perform()
                time.sleep(0.6)
                if combo_preenchido(locator, valor):
                    return True
                print(f"Digitar direto nao selecionou {nome_campo}; procurando na lista aberta: {valor}")
                safe_click(locator)
                time.sleep(0.6)
            if not clicar_opcao_combo(valor, timeout=10):
                print(f"Opcao nao apareceu em {nome_campo}: {valor}")
        except Exception as e:
            print(f"Falha ao pesquisar opcao em {nome_campo} ({valor}): {e}")
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.5)
            continue
        time.sleep(1)
        if combo_preenchido(locator, valor):
            return True
        print(f"{nome_campo} nao foi selecionado (tentativa {tentativa}/{tentativas}): {valor}")
    return False

def primeiro_elemento_visivel(locators, timeout=20):
    fim = time.time() + timeout
    while time.time() < fim:
        for locator in locators:
            for el in driver.find_elements(*locator):
                try:
                    if el.is_displayed() and el.is_enabled():
                        return el
                except Exception:
                    pass
        time.sleep(0.2)
    raise TimeoutException(f"Nenhum elemento visivel encontrado: {locators}")

def preencher_input_visivel(locators, valor, nome_campo, tentativas=3):
    valor_digits = re.sub(r'\D', '', str(valor))
    for tentativa in range(1, tentativas + 1):
        el = primeiro_elemento_visivel(locators, timeout=10)
        scroll_into_view(el)
        try:
            el.click()
            ActionChains(driver)\
                .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)\
                .send_keys(Keys.DELETE)\
                .send_keys(valor)\
                .perform()
        except Exception:
            driver.execute_script(
                "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input', {bubbles: true})); arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
                el,
                valor,
            )
        time.sleep(0.3)
        valor_atual = re.sub(r'\D', '', el.get_attribute('value') or '')
        if valor_digits and valor_digits in valor_atual:
            return True
        print(f"{nome_campo} nao recebeu valor (tentativa {tentativa}/{tentativas}). Valor atual: {valor_atual}")
    return False

def clicar_buscar_fac():
    buscar_locators = [
        (By.XPATH, "//button[contains(normalize-space(.), 'Buscar')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Buscar')]"),
        (By.XPATH, "//*[self::button or self::a][contains(@title, 'Buscar')]"),
        (By.XPATH, "//*[self::button or self::a][contains(@class, 'btn') and .//*[contains(@class, 'search')]]"),
        (By.XPATH, "/html/body/section/section/div/div/div[2]/div/div[1]/form/div[4]/div/button[2]"),
    ]
    el = primeiro_elemento_visivel(buscar_locators, timeout=10)
    scroll_into_view(el)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)

def resumo_resultado_busca():
    texto_pagina = (driver.find_element(By.TAG_NAME, 'body').text or '').strip()
    linhas = driver.find_elements(By.XPATH, "//div[contains(., 'RESULTADO DA BUSCA')]/following::table[1]//tbody/tr")
    linhas_texto = []
    for linha in linhas:
        texto = (linha.text or '').strip()
        if texto:
            linhas_texto.append(texto)
    return texto_pagina, linhas_texto

def aguardar_resultado_busca(timeout=6):
    fim = time.time() + timeout
    texto_pagina = ''
    linhas_texto = []
    while time.time() < fim:
        texto_pagina, linhas_texto = resumo_resultado_busca()
        pagina_upper = texto_pagina.upper()
        tem_contador = re.search(r'EXIBINDO\s+ITENS', pagina_upper) is not None
        if linhas_texto or 'NENHUM' in pagina_upper or tem_contador:
            return texto_pagina, linhas_texto
        time.sleep(0.5)
    return texto_pagina, linhas_texto

def telefone_existe_no_sigavi(telefone, tentativas=3):
    """Confirma na busca do Fac se o telefone realmente entrou no Sigavi."""
    telefone_busca_locators = [
        (By.XPATH, "//input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'telefone')]"),
        (By.XPATH, "//input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'telefone')]"),
        (By.XPATH, "//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'telefone')]"),
        (By.XPATH, "//label[contains(normalize-space(.), 'Telefone')]/following::input[1]"),
        (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[1]/form/div[3]/div/div[3]/input'),
        (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[1]/form/div[2]/div/div/div/div[10]/div[4]/input'),
    ]
    telefone_digits = re.sub(r'\D', '', str(telefone))
    for tentativa in range(1, tentativas + 1):
        try:
            driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
            time.sleep(1)
            if sessao_caiu_para_login():
                relogar_sigavi()
                driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
                time.sleep(1)
            if not preencher_input_visivel(telefone_busca_locators, telefone, 'Telefone da verificacao'):
                continue
            clicar_buscar_fac()
            time.sleep(1)
            texto_pagina, linhas_texto = aguardar_resultado_busca(timeout=8)
            texto_linha = ' | '.join(linhas_texto)
            pagina_upper = texto_pagina.upper()
            telefone_na_linha = telefone_digits and telefone_digits in re.sub(r'\D', '', texto_linha)
            tem_itens = re.search(r'EXIBINDO\s+ITENS\s+1\s*-\s*\d+\s+DE\s+[1-9]\d*', pagina_upper) is not None
            if telefone_na_linha or tem_itens:
                return True, texto_linha or 'Busca por telefone retornou item'
            print(f"Verificacao pos-salvar nao encontrou {telefone} (tentativa {tentativa}/{tentativas}).")
        except Exception as e:
            print(f"Falha na verificacao pos-salvar de {telefone} (tentativa {tentativa}/{tentativas}): {e}")
    return False, 'Telefone nao apareceu na busca apos salvar'

# =========================
# BUSCA POR EMAIL VIA REQUESTS (sem abrir nova aba)
# =========================
def criar_session_requests():
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": f"{BASE_URL}/CRM/fac",
        "Origin": BASE_URL,
        "Accept-Language": "pt-BR,pt;q=0.9",
    })
    for cookie in driver.get_cookies():
        session.cookies.set(cookie['name'], cookie['value'])
    return session

def obter_csrf_token(session):
    try:
        resp = session.get(f"{BASE_URL}/CRM/Fac", timeout=20)
        m = re.search(r'<input[^>]+name="__RequestVerificationToken"[^>]+value="([^"]+)"', resp.text)
        if m:
            return m.group(1)
    except Exception as e:
        print(f"Erro ao obter CSRF token: {e}")
    return None

def _e_fone_valido(d):
    if len(d) not in (10, 11):
        return False
    ddd = int(d[:2])
    if ddd < 11 or ddd > 99:
        return False
    if len(d) == 11 and d[2] != '9':
        return False
    return True

def _extrair_fone_de_texto(texto):
    d = re.sub(r'\D', '', texto)
    if d.startswith('55') and len(d) >= 12:
        d = d[2:]
    if _e_fone_valido(d):
        return d
    # tenta concatenar blocos de dígitos adjacentes (ex: "9 9999-9999" com DDD separado)
    blocos = re.findall(r'\d+', texto)
    for i in range(len(blocos)):
        acum = ''
        for j in range(i, min(i + 5, len(blocos))):
            acum += blocos[j]
            if len(acum) > 11:
                break
            if len(acum) >= 10:
                d2 = acum
                if d2.startswith('55') and len(d2) >= 12:
                    d2 = d2[2:]
                if _e_fone_valido(d2):
                    return d2
    return None

def extrair_telefone_do_html(html):
    """Analisa a resposta da busca. Retorna (telefone|None, classificacao).

    classificacao:
      'encontrado'   -> achou telefone valido
      'sem_resultado'-> ausencia confirmada (texto de "sem resultado" ou grade vazia)
      'suspeito'     -> resposta anomala/incompleta (provavel throttle do Sigavi);
                        deve ser retentada para nao virar falso "nao encontrado"
    """
    html_lower = html.lower()
    tem_kw_sem_resultado = any(k in html_lower for k in (
        'não retornou', 'nao retornou', 'sem resultado', 'nenhum registro',
        'nao foram encontrados', 'não foram encontrados',
    ))

    # respostas curtas (< 5000 chars) só são "sem resultado" quando trazem a
    # confirmação textual; sem ela, uma resposta curta é anômala (throttle)
    if len(html) < 5000 and tem_kw_sem_resultado:
        return None, 'sem_resultado'

    # 1) células marcadas como dados pessoais (telefone preenchido)
    dados_pessoais = re.findall(r'<td[^>]*data-dados-pessoais="true"[^>]*>([^<]*)</td>', html)
    for celula in dados_pessoais:
        tel = _extrair_fone_de_texto(celula)
        if tel:
            return tel, 'encontrado'

    # 2) todas as <td> do tbody
    tbody = re.search(r'<tbody>(.*?)</tbody>', html, re.DOTALL)
    if tbody:
        for celula in re.findall(r'<td[^>]*>([^<]*)</td>', tbody.group(1)):
            tel = _extrair_fone_de_texto(celula)
            if tel:
                return tel, 'encontrado'
        # veio a grade de resultados, mas sem telefone util -> ausencia real
        return None, 'sem_resultado'

    # confirmacao textual de ausencia mesmo sem grade
    if tem_kw_sem_resultado:
        return None, 'sem_resultado'

    # sem grade e sem confirmacao textual -> resposta anomala (provavel throttle)
    return None, 'suspeito'

def buscar_telefone_por_email(session, csrf_token, email):
    """Busca o telefone de um lead pelo email. Retorna (telefone|None, erro|None).

    Seguro para uso em paralelo: a renovacao de sessao expirada acontece sob
    _token_lock, ja que mexe nos cookies do driver Selenium (nao thread-safe).
    """
    global req_csrf_token
    payload = {
        'FacBusca': 'true',
        '__RequestVerificationToken': csrf_token,
        'Numero': '', 'Fase0': 'false', 'Fase1': 'false', 'Fase2': 'false',
        'Fase3': 'false', 'Fase4': 'false', 'Fase5': 'false',
        'RetornoVisita': 'false', 'FaseAnalise': 'false',
        'Fase6': 'false', 'Fase7': 'false',
        'Transferida': 'false', 'SemCorretor': 'false', 'SemPerfilInteresse': 'false',
        'CadastroDe': '', 'CadastroAte': '', 'AtualizacaoDe': '', 'AtualizacaoAte': '',
        'ValidadeDe': '', 'ValidadeAte': '', 'TransferenciaDe': '', 'TransferenciaAte': '',
        'StatusGestao': '', 'IdMidia': '0', 'IdSituacaoAtendimento': '0',
        'TipoReferencia': '', 'TipoOrigem': '', 'IdClassificacao': '0',
        'IdCargo': '0', 'IdMotivoFinalizacao': '0', 'IdCentralAtendimento': '0',
        'EmpreendimentoUsados': '', 'TarefaAgendada': '', 'IdAgencia': '',
        'EquipeOrigem': 'false', 'CorretorOrigem': 'false', 'ParceriaInterna': 'false',
        'EquipeGerente2': '', 'Cliente': '', 'Email': email,
        'DDI': '', 'Telefone': '', 'CpfCnpj': '',
        'Compra': 'false', 'Locacao': 'false', 'IdFinalidade': '0',
        'Dormitorio': '0,10', 'Suite': '0,10', 'Vaga': '0,10',
        'AreaTotal': '50,500', 'ValorDe': '0', 'ValorAte': '0',
        'CidadeInteresse': '', 'IdZona': '', 'IdImovelClassificacao': '',
        'IdProfissao': '0', 'Renda': '1000,20000',
        'NascimentoDe': '', 'NascimentoAte': '', 'IdFaixaEtaria': '0',
        'IdSexo': '0', 'IdEstadoCivil': '0', 'Filhos': '',
        'CidadeMora': '', 'EmpresaTrabalha': '', 'CidadeTrabalha': '',
        'IncluirHistoricoEmpreendimento': 'false', 'CentralAtendimentoHistorico': 'false',
        'X-Requested-With': 'XMLHttpRequest',
    }
    def _sessao_expirada(r):
        return 'ReturnUrl' in r.url or 'Login' in r.url or 'login' in r.text[:500].lower()

    headers = {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'Accept': '*/*'}
    ultimo_erro = None

    # Retentamos respostas suspeitas/erros: o Sigavi devolve respostas vazias
    # quando esta sob carga, e isso nao pode virar falso "nao encontrado".
    for tentativa in range(1, CONSULTA_TENTATIVAS + 1):
        try:
            token_atual = req_csrf_token or csrf_token
            resp = session.post(
                f"{BASE_URL}/CRM/Fac/Busca",
                data={**payload, '__RequestVerificationToken': token_atual},
                headers=headers,
                timeout=30,
            )
            if resp.status_code == 200:
                # sessão expirada → servidor retorna página de login
                if _sessao_expirada(resp):
                    # renova sob lock: driver.get_cookies() nao e thread-safe
                    with _token_lock:
                        for cookie in driver.get_cookies():
                            session.cookies.set(cookie['name'], cookie['value'])
                        novo_token = obter_csrf_token(session)
                        if novo_token:
                            req_csrf_token = novo_token
                    token_atual = req_csrf_token
                    if not token_atual:
                        ultimo_erro = 'Sessao expirada e CSRF token nao foi renovado.'
                        time.sleep(CONSULTA_BACKOFF * tentativa)
                        continue
                    resp = session.post(
                        f"{BASE_URL}/CRM/Fac/Busca",
                        data={**payload, '__RequestVerificationToken': token_atual},
                        headers=headers,
                        timeout=30,
                    )
                    if _sessao_expirada(resp):
                        ultimo_erro = 'Sessao expirada mesmo apos renovar cookies.'
                        time.sleep(CONSULTA_BACKOFF * tentativa)
                        continue

                tel, classificacao = extrair_telefone_do_html(resp.text)
                if classificacao == 'encontrado':
                    return tel, None
                if classificacao == 'sem_resultado':
                    return None, None  # ausencia confiavel
                # 'suspeito' -> provavel throttle; espera e tenta de novo
                ultimo_erro = f'Resposta suspeita ({len(resp.text)} chars) - possivel limite do Sigavi.'
            else:
                ultimo_erro = f'Status HTTP inesperado: {resp.status_code}'
        except Exception as e:
            ultimo_erro = str(e)

        time.sleep(CONSULTA_BACKOFF * tentativa)

    return None, ultimo_erro

def normalizar_texto(valor: str) -> str:
    if valor is None:
        return ''
    nfkd = unicodedata.normalize('NFKD', str(valor))
    sem_acento = ''.join(c for c in nfkd if not unicodedata.combining(c))
    upper = sem_acento.upper()
    apenas_alfa = re.sub(r'[^A-Z0-9]+', ' ', upper)
    return re.sub(r'\s+', ' ', apenas_alfa).strip()

def selecionar_midia_por_posicao(posicao: int):
    ac = ActionChains(driver)
    for _ in range(posicao):
        ac.send_keys(Keys.ARROW_DOWN)
    ac.send_keys(Keys.ENTER)
    ac.perform()

midia_mapeamentos_brutos = [
    ("123I", 1),
    ("APTO.VC", 2),
    ("CARTEIRA", 3),
    ("CHAVES NA MAO", 4),
    ("CHAVES NA MÃO", 4),
    ("FACEBOOK", 5),
    ("FORMULARIO GOOGLE", 6),
    ("FORMULÁRIO GOOGLE", 6),
    ("GOOGLE DISPLAY", 7),
    ("GOOGLE SEARCH", 8),
    ("IMOVEL WEB", 9),
    ("IMÓVEL WEB", 9),
    ("INCORPORADOR", 11),
    ("INDICACAO", 12),
    ("INDICAÇÃO", 12),
    ("IND. CORRETOR", 12),
    ("IND CORRETOR", 12),
    ("INSTAGRAM", 13),
    ("LINKEDIN", 14),
    ("OLX", 15),
    ("OUTROS", 16),
    ("PADARIA", 17),
    ("PLACA", 18),
    ("RD STATION", 19),
    ("RETORNO", 20),
    ("SITE", 21),
    ("VISITACAO", 21),
    ("VISITAÇÃO", 21),
    ("STAND", 21),
    ("ACAO DE RUA", 21),
    ("AÇÃO DE RUA", 21),
    ("STAND/ACAO DE RUA", 21),
    ("STAND/AÇÃO DE RUA", 21),
    ("TWITTER", 23),
    ("VIVA REAL", 24),
    ("VIZINHO", 25),
    ("WHATSAPP", 26),
    ("YOUTUBE", 27),
    ("ZAP", 28),
    # Base Cora
    ("VISITA DIRETA AO STANDE", 21),
    ("SITE FORMULARIO", 21),
    ("TELEFONE", 16),
    ("CORA PINHEIROS", 16),
]

midia_por_tipo_plantao = {}
for chave, posicao in midia_mapeamentos_brutos:
    chave_norm = normalizar_texto(chave)
    if chave_norm not in midia_por_tipo_plantao:
        midia_por_tipo_plantao[chave_norm] = posicao

canal_plantao_tipos = {
    normalizar_texto("VISITACAO"),
    normalizar_texto("VISITAÇÃO"),
    normalizar_texto("RETORNO"),
    normalizar_texto("VISITA DIRETA AO STANDE"),
}

canal_carteira_tipos = {                                        
    normalizar_texto("INDICACAO"),
    normalizar_texto("INDICAÇÃO"),
    normalizar_texto("IND. CORRETOR"),
    normalizar_texto("IND CORRETOR"),
    normalizar_texto("CARTEIRA"),
}

# =========================
# LOGIN
# =========================
driver.get("https://abyara.sigavi360.com.br/Acesso/Login?ReturnUrl=%2F")
time.sleep(2)

wait_visible((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[1]/div/input"))\
    .send_keys(SIGAVI_LOGIN)
wait_visible((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[2]/div/input"))\
    .send_keys(SIGAVI_SENHA)
safe_click((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[3]/div/button"))

# =========================
# VALIDA O LOGIN (melhoria nº 8)
# Antes a automacao seguia mesmo se o login falhasse, gerando erro em TODAS as
# linhas (foi o bug dos 51 erros). Agora confirma que de fato saiu da pagina de
# login; se nao sair, para na hora com aviso claro em vez de rodar tudo errado.
# =========================
# Da tempo do POST de login processar e verifica de forma robusta: abre a home;
# se a sessao nao estiver autenticada, o Sigavi redireciona de volta pro login.
time.sleep(3)
driver.get(BASE_URL + "/")
time.sleep(2)
_url_pos_login = driver.current_url.lower()
if "login" in _url_pos_login or "/acesso" in _url_pos_login:
    print(
        "\nERRO: nao foi possivel entrar no Sigavi — o login nao foi concluido.\n"
        "Causas mais comuns: usuario/senha incorretos, outra sessao do Sigavi "
        "aberta em outro lugar (sessao concorrente) ou captcha.\n"
        "Confira as credenciais, feche outras sessoes do Sigavi e rode de novo.\n"
    )
    try:
        driver.quit()
    except Exception:
        pass
    sys.exit(1)
print("Login no Sigavi confirmado.")

# =========================
# SESSION REQUESTS (somente consulta)
# =========================
req_session = None
req_csrf_token = None
if MODE == 'consulta':
    req_session = criar_session_requests()
    req_csrf_token = obter_csrf_token(req_session)
    if req_csrf_token:
        print("Session de busca por email pronta.")
    else:
        print("AVISO: Não foi possível obter CSRF token. Busca por email desabilitada.")
        req_session = None

# detecta coluna de email na planilha
_email_col = next(
    (c for c in df.columns if re.match(r'e.?mail', c, re.IGNORECASE)),
    None
)
if _email_col:
    print(f"Coluna de email detectada: '{_email_col}'")
else:
    print("AVISO: Nenhuma coluna de email encontrada na planilha.")

# =========================
# LOOP DE CADASTRO
# =========================
_ultimo_index, resultados_email, resultados_cadastro = carregar_progresso()

# Leads cujo CORRETOR DE ORIGEM nao existe no corretores.json (cadastrados como
# Tabatanascimento/Corretor Inativo) — relatorio proprio, como na automacao do
# Pedro, pra Marketing revisar depois quem caiu nessa regra.
resultados_corretor_inativo = []

if MODE == 'consulta':
    print("Modo selecionado: somente consulta.")
else:
    print("Modo selecionado: somente cadastro.")
    driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
    time.sleep(2)

# mapa para tolerar variações de caixa no dicionário
mapa_corretores = {str(k).upper(): v for k, v in corretores_gerentes.items()}

# Nomes oficiais das equipes (valores do corretores.json). Usados pra traduzir
# apelido vindo da planilha (ex.: 'ELAINE') pro nome que existe no combo do
# Sigavi (ex.: 'Elainemaion') antes de tentar selecionar.
equipes_oficiais = sorted(set(corretores_gerentes.values()))

# mapa de fallback: só o nome base antes de sufixos como "- Inc", "- Pagadoria", "- IND", etc.
# ex: "DIMI - INC" → chave base "DIMI"
mapa_corretores_base = {}
for k, v in mapa_corretores.items():
    base = re.split(r'\s*[-|]\s*', k)[0].strip()
    if base and base not in mapa_corretores_base:
        mapa_corretores_base[base] = v

def _salvar_excel_resultado(anunciar=False):
    os.makedirs(RESULT_DIR, exist_ok=True)
    arquivos_gerados = []

    if MODE == 'consulta':
        colunas = ['Linha', 'Nome', 'Email', 'Telefone', 'Status', 'Detalhe']
        df_todos = pd.DataFrame(resultados_email, columns=colunas)
        if df_todos.empty:
            df_todos = pd.DataFrame(columns=colunas)

        df_encontrados = df_todos[df_todos['Status'] == 'encontrado'].reset_index(drop=True)
        df_nao_encontrados = df_todos[df_todos['Status'] == 'nao_encontrado'].reset_index(drop=True)
        df_erros = df_todos[df_todos['Status'] == 'erro_consulta'].reset_index(drop=True)

        relatorios = {
            'encontrados': df_encontrados,
            'nao_encontrados': df_nao_encontrados,
            'erros_consulta': df_erros,
        }
    else:
        colunas = ['Linha', 'Nome', 'Email', 'Telefone', 'Status', 'Detalhe']
        df_todos = pd.DataFrame(resultados_cadastro, columns=colunas)
        if df_todos.empty:
            df_todos = pd.DataFrame(columns=colunas)

        df_inativos = pd.DataFrame(resultados_corretor_inativo, columns=colunas)
        if df_inativos.empty:
            df_inativos = pd.DataFrame(columns=colunas)

        relatorios = {
            'cadastrados': df_todos[df_todos['Status'] == 'cadastrado'].reset_index(drop=True),
            'duplicados': df_todos[df_todos['Status'] == 'duplicado'].reset_index(drop=True),
            'nao_cadastrados': df_todos[df_todos['Status'] == 'nao_cadastrado'].reset_index(drop=True),
            'erros_cadastro': df_todos[df_todos['Status'] == 'erro_cadastro'].reset_index(drop=True),
            'corretores_inativos': df_inativos.reset_index(drop=True),
        }

    for sufixo, dataframe in relatorios.items():
        arquivo = _resultado_file(sufixo)
        dataframe.to_excel(arquivo, sheet_name=sufixo[:31], index=False)
        arquivos_gerados.append(arquivo)

    if anunciar:
        for arquivo in arquivos_gerados:
            print(f"RESULT_FILE={arquivo}")
        if MODE == 'consulta':
            print(
                "Resumo consulta: "
                f"{len(relatorios['encontrados'])} encontrado(s), "
                f"{len(relatorios['nao_encontrados'])} nao encontrado(s), "
                f"{len(relatorios['erros_consulta'])} erro(s)."
            )
        else:
            print(
                "Resumo cadastro: "
                f"{len(relatorios['cadastrados'])} cadastrado(s), "
                f"{len(relatorios['duplicados'])} duplicado(s), "
                f"{len(relatorios['nao_cadastrados'])} nao cadastrado(s), "
                f"{len(relatorios['erros_cadastro'])} erro(s), "
                f"{len(relatorios['corretores_inativos'])} com corretor inativo."
            )


def salvar_estado(ultimo_index, anunciar=False):
    salvar_progresso(ultimo_index, resultados_email, resultados_cadastro)
    _salvar_excel_resultado(anunciar=anunciar)


def emitir_progresso():
    """Emite uma linha estruturada com os contadores REAIS da execucao.

    A interface le essa linha (prefixo PROGRESS=) para mostrar os numeros e a
    barra de progresso. Nao depende dos logs visiveis, entao nao 'esquece'
    sucessos antigos quando o buffer de log rotaciona.
    """
    total = len(df)
    if MODE == 'consulta':
        sucessos = sum(1 for r in resultados_email if r['Status'] == 'encontrado')
        pendentes_ct = sum(1 for r in resultados_email if r['Status'] == 'nao_encontrado')
        erros = sum(1 for r in resultados_email if r['Status'] == 'erro_consulta')
        processados = len(resultados_email)
    else:
        sucessos = sum(1 for r in resultados_cadastro if r['Status'] == 'cadastrado')
        pendentes_ct = sum(
            1 for r in resultados_cadastro
            if r['Status'] in ('duplicado', 'nao_cadastrado')
        )
        erros = sum(1 for r in resultados_cadastro if r['Status'] == 'erro_cadastro')
        processados = len(resultados_cadastro)

    print("PROGRESS=" + json.dumps({
        'processados': processados,
        'total': total,
        'sucessos': sucessos,
        'pendentes': pendentes_ct,
        'erros': erros,
    }), flush=True)


def _processar_linha_consulta(index, row):
    """Resolve o telefone de uma unica linha (sem efeitos colaterais de estado).

    Roda em paralelo dentro do ThreadPoolExecutor: faz a busca HTTP quando
    necessario e devolve o dict pronto para a planilha de resultado.
    """
    nome = str(row.get('NOME') or '').strip()
    email_raw = str(row.get(_email_col) or '').strip() if _email_col else ''
    telefone = re.sub(r'\D', '', str(row.get('FONE2') or ''))

    def _resultado(status, telefone_final, detalhe):
        return {
            'Linha': index + 1,
            'Nome': nome,
            'Email': email_raw,
            'Telefone': telefone_final,
            'Status': status,
            'Detalhe': detalhe,
        }

    if not email_raw:
        return _resultado('erro_consulta', '', 'Email ausente na planilha.')
    if len(telefone) >= 11:
        return _resultado('encontrado', telefone, 'Telefone ja estava na planilha.')
    if not (req_session and req_csrf_token):
        return _resultado('erro_consulta', '', 'Sessao de consulta indisponivel.')

    tel, erro = buscar_telefone_por_email(req_session, req_csrf_token, email_raw)
    tel = tel or ''
    if len(tel) >= 10:
        return _resultado('encontrado', tel, 'Telefone encontrado por email.')
    if erro:
        return _resultado('erro_consulta', '', erro)
    return _resultado('nao_encontrado', '', 'Telefone nao encontrado por email.')


if MODE == 'consulta':
    ultimo_processado = _ultimo_index
    pendentes = [(index, row) for index, row in df.iterrows() if index > _ultimo_index]
    # blocos: paraleliza a rede dentro do bloco e faz checkpoint ao fim de cada um
    bloco_tam = max(AUTOSAVE_EVERY, CONSULTA_WORKERS)
    total_pendentes = len(pendentes)
    processados = 0
    marcas = {'encontrado': '✓', 'nao_encontrado': '✗', 'erro_consulta': '!'}
    print(f"Consulta paralela: {total_pendentes} linha(s) a processar, {CONSULTA_WORKERS} em paralelo.")
    emitir_progresso()

    try:
        for inicio in range(0, total_pendentes, bloco_tam):
            if parada_solicitada():
                print("\nParada solicitada pela interface. Salvando resultados...")
                salvar_estado(ultimo_processado, anunciar=True)
                emitir_progresso()
                driver.quit()
                raise SystemExit(0)

            bloco = pendentes[inicio:inicio + bloco_tam]
            resultados_bloco = {}
            with ThreadPoolExecutor(max_workers=CONSULTA_WORKERS) as executor:
                futuros = {
                    executor.submit(_processar_linha_consulta, index, row): index
                    for index, row in bloco
                }
                for futuro in as_completed(futuros):
                    resultados_bloco[futuros[futuro]] = futuro.result()

            # consolida na ordem original do bloco (resultado estavel e retomavel)
            for index, _row in bloco:
                resultado = resultados_bloco[index]
                resultados_email.append(resultado)
                processados += 1
                marca = marcas.get(resultado['Status'], '?')
                print(f"[{processados}/{total_pendentes}] linha {resultado['Linha']} "
                      f"{resultado['Email']} [{marca}] {resultado['Telefone']}")
                ultimo_processado = index

            salvar_estado(ultimo_processado)
            emitir_progresso()

        print("Consulta concluida!")
        salvar_estado(ultimo_processado, anunciar=True)
        emitir_progresso()
        driver.quit()
        raise SystemExit(0)

    except KeyboardInterrupt:
        print("\n\nPausado pelo usuario. Salvando resultados...")
        salvar_estado(ultimo_processado, anunciar=True)
        emitir_progresso()
        driver.quit()
        raise SystemExit(0)

try:
    for index, row in df.iterrows():
        if index <= _ultimo_index:
            continue
        if parada_solicitada():
            print("\nParada solicitada pela interface. Salvando resultados...")
            salvar_estado(index - 1, anunciar=True)
            emitir_progresso()
            driver.quit()
            raise SystemExit(0)

        emitir_progresso()
        nome = str(row.get('NOME') or '').strip()
        email_raw = ''
        if _email_col:
            email_raw = str(row.get(_email_col) or '').strip()

        # sanitiza telefone (só dígitos)
        telefone_raw = str(row.get('FONE2') or '')
        telefone = re.sub(r'\D', '', telefone_raw)
    
        if len(telefone) < 11:
            print(f"[NAO CADASTRADO] linha {index + 1}: telefone ausente ou invalido.")
            resultados_cadastro.append({
                'Linha': index + 1,
                'Nome': nome,
                'Email': email_raw,
                'Telefone': telefone,
                'Status': 'nao_cadastrado',
                'Detalhe': 'Telefone ausente ou invalido na planilha.',
            })
            salvar_estado(index)
            continue
    
        corretor_original_raw = str(row.get('CORRETOR DE ORIGEM') or '')
        corretor_original_norm = re.sub(r'\s+', ' ', corretor_original_raw).strip().upper()
    
        tipo_plantao_raw = str(row.get('TIPO PLANTAO') or '').strip()
        tipo_plantao_norm = normalizar_texto(tipo_plantao_raw)
        posicao_midia = midia_por_tipo_plantao.get(tipo_plantao_norm, 16)  # default: Outros

        # Regra de negocio (igual a automacao do Pedro/automa-o_abyara): a
        # equipe SEMPRE sai do corretores.json a partir do CORRETOR DE ORIGEM
        # — a coluna GERENTE da planilha NAO e usada (vem com apelido, ex.
        # 'ELAINE', que nao existe no combo do Sigavi). Corretor fora do json
        # = corretor inativo -> cadastra como Tabatanascimento/Corretor Inativo
        # (cadastra mesmo assim, nao pula o lead).
        if corretor_original_norm in mapa_corretores:
            gerente  = mapa_corretores[corretor_original_norm]
            corretor = corretor_original_raw.strip()
        elif corretor_original_norm in mapa_corretores_base:
            gerente  = mapa_corretores_base[corretor_original_norm]
            corretor = corretor_original_raw.strip()
            print(f"Corretor '{corretor_original_raw}' encontrado via nome base (sem sufixo). Gerente: {gerente}")
        else:
            print(f"Corretor '{corretor_original_raw}' nao encontrado no corretores.json (inativo). Cadastrando como equipe Tabatanascimento / Corretor Inativo.")
            gerente  = "Tabatanascimento"
            corretor = "Corretor Inativo"
            resultados_corretor_inativo.append({
                'Linha': index + 1,
                'Nome': nome,
                'Email': email_raw,
                'Telefone': telefone,
                'Status': 'corretor_inativo',
                'Detalhe': f"Corretor '{corretor_original_raw.strip() or '(vazio)'}' nao encontrado no corretores.json; cadastrado como Tabatanascimento/Corretor Inativo.",
            })

        # Traduz apelido de equipe pro nome oficial quando bater com exatamente
        # UMA equipe do corretores.json (ex.: ELAINE -> Elainemaion). Sem isso o
        # combo do Sigavi nao acha a opcao e o lead e pulado a toa.
        gerente_norm_aj = normalizar_texto(gerente)
        if gerente_norm_aj and all(normalizar_texto(e) != gerente_norm_aj for e in equipes_oficiais):
            candidatas = [
                e for e in equipes_oficiais
                if gerente_norm_aj in normalizar_texto(e) or normalizar_texto(e) in gerente_norm_aj
            ]
            if len(candidatas) == 1:
                print(f"Equipe '{gerente}' ajustada para '{candidatas[0]}' (nome oficial do corretores.json).")
                gerente = candidatas[0]
            elif len(candidatas) > 1:
                print(f"Equipe '{gerente}' bate com varias do corretores.json ({', '.join(candidatas[:4])}); mantendo como veio.")
    
        canal_setas = 4 if tipo_plantao_norm in canal_plantao_tipos else 1
        if tipo_plantao_norm in canal_carteira_tipos:
            canal_setas = 1
    
    
        # Cada lead tem ate 2 tentativas completas: se falhar no meio (combo,
        # modal, crash do browser), roda o lead de novo uma vez; persistindo o
        # erro, registra na planilha e PULA pro proximo (nao trava o job).
        for tentativa_lead in (1, 2):
          try:
            # Página de busca (apenas preenche telefone; não valida duplicidade aqui)
            telefone_busca_locator = (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[1]/form/div[2]/div/div/div/div[10]/div[4]/input')
            telefone_elem_busca = None
            for tentativa in range(3):
                if parada_solicitada():
                    raise KeyboardInterrupt
                driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
                pausa_cadastro(3)
                try:
                    telefone_elem_busca = wait_visible(telefone_busca_locator, timeout=20)
                    break
                except TimeoutException:
                    # Se o Sigavi mandou de volta pro login, a sessao caiu:
                    # reloga e tenta de novo em vez de queimar as 3 tentativas
                    # na tela de login (onde o campo de telefone nunca existe).
                    if sessao_caiu_para_login():
                        print(f"Página /CRM/Fac redirecionou pro login — sessão expirou (tentativa {tentativa+1}/3). Relogando...")
                        try:
                            relogar_sigavi()
                        except Exception as exc_relogin:
                            print(f"Relogin falhou: {exc_relogin}")
                        continue
                    try:
                        url_atual = driver.current_url
                    except WebDriverException:
                        url_atual = '?'
                    print(f"Página /CRM/Fac não carregou (tentativa {tentativa+1}/3, url atual: {url_atual}). Tentando novamente...")
            if telefone_elem_busca is None:
                print(f"Não foi possível carregar a página de busca após 3 tentativas. Pulando {nome}.")
                resultados_cadastro.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'erro_cadastro',
                    'Detalhe': 'Pagina /CRM/Fac nao carregou para verificar duplicidade.',
                })
                salvar_estado(index)
                break
            scroll_into_view(telefone_elem_busca)
    
            # limpa e digita
            ActionChains(driver)\
                .click(on_element=telefone_elem_busca)\
                .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)\
                .send_keys(Keys.DELETE)\
                .send_keys(telefone)\
                .perform()
    
            # dispara busca e verifica duplicidade antes de ir para cadastro
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            pausa_cadastro(2.5)  # aguarda carregamento da grade de resultados
            resultado_busca_locator = (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[2]/div[2]/div[2]/div/div[4]/table/tbody/tr')
            duplicado = False
            try:
                linhas = WebDriverWait(driver, 6).until(EC.presence_of_all_elements_located(resultado_busca_locator))
                if linhas:
                    primeira_td = linhas[0].find_element(By.XPATH, './td[1]')
                    texto = (primeira_td.text or '').strip()
                    texto_linha = (linhas[0].text or '').strip()
                    if (texto and not texto.upper().startswith('NENHUM')) or (texto_linha and not texto_linha.upper().startswith('NENHUM')):
                        print(f"[DUPLICADO] {telefone}")
                        duplicado = True
            except (TimeoutException, Exception):
                pass
    
            if duplicado:
                resultados_cadastro.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'duplicado',
                    'Detalhe': 'Telefone ja encontrado no Sigavi antes do cadastro.',
                })
                salvar_estado(index)
                break
    
            # navegação leve (como no seu script)
            ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
            pausa_cadastro(0.5)
            ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
            pausa_cadastro(0.5)

            # Vai direto ao cadastro
            driver.get('https://abyara.sigavi360.com.br/CRM/Fac/Cadastro')
            pausa_cadastro(2)

            # === BLOCO CADASTRO ===
            wait_visible((By.ID, 'Nome')).send_keys(nome)
            pausa_cadastro(2.5)

            # Abre o bloco de telefones
            safe_click((By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/div/a'))
            pausa_cadastro(1)
    
            # Seleciona "Celular" no tipo (setas + enter)
            celular_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[1]/span[1]/span/span[1]')
            safe_click(celular_combo_locator)
            ActionChains(driver).send_keys(Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ENTER).perform()
    
            # Preenche número
            telefone_grid_input_locator = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[3]/input')
            tel_input = wait_visible(telefone_grid_input_locator)
            tel_input.click()
            tel_input.send_keys(telefone)
            pausa_cadastro(1)

            # Adiciona telefone (ícone de +/confirmar)
            safe_click((By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[4]/a[1]/span'))
            pausa_cadastro(1)
    
            # Canal (SMS)
            sms_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[1]/span[2]/span/span[1]')
            safe_click(sms_combo_locator)
            ac_canal = ActionChains(driver)
            for _ in range(canal_setas):
                ac_canal.send_keys(Keys.ARROW_DOWN)
            ac_canal.send_keys(Keys.ENTER).perform()
            pausa_cadastro(1)

            # Mídia
            midia_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[2]/span[2]/span/span[1]')
            safe_click(midia_combo_locator)
            # seleciona conforme TIPO PLANTAO
            selecionar_midia_por_posicao(posicao_midia)
            pausa_cadastro(1)

            # Equipe (gerente) — seleciona digitando e CONFERE se realmente pegou.
            # Se a equipe/corretor estiverem no corretores.json mas NAO aparecerem
            # no combo do Sigavi (ex.: JOTACE/Logan em 2026-06-11), cai pro
            # fallback Tabatanascimento/Corretor Inativo em vez de pular o lead
            # — mesma regra de quem nem esta no json.
            equipe_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[1]/span[1]/span/span[1]')
            corretor_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/span[1]/span/span[1]')
            ja_eh_inativo = (
                normalizar_texto(gerente) == normalizar_texto('Tabatanascimento')
                and normalizar_texto(corretor) == normalizar_texto('Corretor Inativo')
            )

            # Uma tratativa so pro nome da planilha: a tentativa unica ja digita
            # E procura na lista aberta; se nao achou, vai direto pro fallback
            # de Corretor Inativo em vez de repetir tudo 4x (lento a toa).
            # So o fallback (Tabatanascimento/Corretor Inativo) mantem as 4
            # tentativas, porque falhar nele significa pular o lead.
            equipe_ok = selecionar_combo_texto(
                equipe_combo_locator, gerente, 'Equipe',
                tentativas=4 if ja_eh_inativo else 1,
            )
            if not equipe_ok and not ja_eh_inativo:
                print(f"Equipe '{gerente}' nao apareceu no combo do Sigavi. Cadastrando {nome} como Tabatanascimento / Corretor Inativo.")
                resultados_corretor_inativo.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'corretor_inativo',
                    'Detalhe': f"Equipe '{gerente}' nao apareceu no combo do Sigavi; cadastrado como Tabatanascimento/Corretor Inativo.",
                })
                gerente, corretor = 'Tabatanascimento', 'Corretor Inativo'
                ja_eh_inativo = True
                equipe_ok = selecionar_combo_texto(equipe_combo_locator, gerente, 'Equipe')
            if not equipe_ok:
                print(f"[ERRO] Equipe '{gerente}' nao selecionada para {nome}. Pulando lead.")
                resultados_cadastro.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'erro_cadastro',
                    'Detalhe': f"Equipe '{gerente}' nao foi selecionada no formulario.",
                })
                salvar_estado(index)
                break

            # Corretor — espera a lista da equipe carregar antes de selecionar
            print(f"Aguardando lista de corretores da equipe {gerente} para selecionar: {corretor}")
            pausa_cadastro(2)
            corretor_ok = selecionar_combo_texto(
                corretor_combo_locator, corretor, 'Corretor',
                tentativas=4 if ja_eh_inativo else 1,
            )
            if not corretor_ok and not ja_eh_inativo:
                print(f"Corretor '{corretor}' nao apareceu na lista da equipe {gerente}. Cadastrando {nome} como Tabatanascimento / Corretor Inativo.")
                resultados_corretor_inativo.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'corretor_inativo',
                    'Detalhe': f"Corretor '{corretor}' nao apareceu no combo da equipe {gerente} no Sigavi; cadastrado como Tabatanascimento/Corretor Inativo.",
                })
                gerente, corretor = 'Tabatanascimento', 'Corretor Inativo'
                ja_eh_inativo = True
                if selecionar_combo_texto(equipe_combo_locator, gerente, 'Equipe'):
                    pausa_cadastro(2)
                    corretor_ok = selecionar_combo_texto(corretor_combo_locator, corretor, 'Corretor')
            if not corretor_ok:
                print(f"[ERRO] Corretor '{corretor}' nao selecionado para {nome}. Pulando lead.")
                resultados_cadastro.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'erro_cadastro',
                    'Detalhe': f"Corretor '{corretor}' nao foi selecionado no formulario (equipe: {gerente}).",
                })
                salvar_estado(index)
                break
    
            # Abre modal "Imóvel/Origem"
            safe_click((By.XPATH, "/html/body/div[2]/form/div[3]/div/div/div[1]/div[2]/div[2]/a/span"))
            modal_container = wait_visible((By.XPATH, "/html/body/div[2]/div[2]/div/div"))
    
            # Seleciona a opção dentro do modal, SE existir. No modal atual do
            # Sigavi nao tem mais os labels de tipo (so campo + Buscar); esperar
            # 20s por um label fantasma era o que "travava" e virava o falso
            # "Browser caiu" (TimeoutException herda de WebDriverException).
            try:
                opcao_modal = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[1]/div/div/label[2]"))
                )
                opcao_modal.click()
                pausa_cadastro(0.5)
            except TimeoutException:
                print("Modal sem opcoes de tipo (label[2] nao existe); indo direto pro campo de busca.")

            # Preenche o empreendimento e CONFERE que o texto entrou no campo.
            # O campo e localizado pelo placeholder 'Empreendimento' (robusto a
            # mudancas de layout), com o XPath absoluto antigo como fallback.
            empreendimento_input_locators = [
                (By.XPATH, "//input[contains(translate(@placeholder, 'EMPRENDIMTO', 'emprendimto'), 'empreendimento')]"),
                (By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/input"),
                (By.XPATH, "/html/body/div[2]/div[2]/div/div//input[@type='text']"),
            ]
            empreendimento_ok = False
            for tentativa_emp in range(1, 4):
                try:
                    your_code = primeiro_elemento_visivel(empreendimento_input_locators, timeout=8)
                except TimeoutException:
                    print(f"Campo de busca do modal nao apareceu (tentativa {tentativa_emp}/3).")
                    continue
                scroll_into_view(your_code)
                try:
                    # jeito do Pedro: send_keys direto no elemento (foca sozinho)
                    your_code.clear()
                    your_code.send_keys(EMPREENDIMENTO_BUSCA)
                except Exception as e_fill:
                    print(f"send_keys no campo do modal falhou ({e_fill}); preenchendo via JS.")
                    driver.execute_script(
                        "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input', {bubbles: true})); arguments[0].dispatchEvent(new Event('change', {bubbles: true}));",
                        your_code,
                        EMPREENDIMENTO_BUSCA,
                    )
                pausa_cadastro(0.5)
                valor_atual = (your_code.get_attribute('value') or '').strip()
                if EMPREENDIMENTO_BUSCA.lower() in valor_atual.lower():
                    empreendimento_ok = True
                    print(f"Empreendimento '{valor_atual}' digitado no modal.")
                    break
                print(f"Campo Empreendimento nao recebeu '{EMPREENDIMENTO_BUSCA}' (tentativa {tentativa_emp}/3). Valor atual: '{valor_atual}'")
            if not empreendimento_ok:
                raise Exception(f"Campo Empreendimento do modal nao recebeu o texto '{EMPREENDIMENTO_BUSCA}'.")

            # Confirma modal (botao Buscar) — por texto, com fallback no XPath antigo
            buscar_modal_locators = [
                (By.XPATH, "/html/body/div[2]/div[2]/div/div//button[contains(normalize-space(.), 'Buscar')]"),
                (By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/button"),
            ]
            botao_buscar = primeiro_elemento_visivel(buscar_modal_locators, timeout=10)
            scroll_into_view(botao_buscar)
            try:
                botao_buscar.click()
            except Exception:
                driver.execute_script("arguments[0].click();", botao_buscar)
            pausa_cadastro(1)

            # Botão que às vezes fica atrás de overlay
            safe_click((By.CSS_SELECTOR, "#dvImovelOrigemComando a"))
            pausa_cadastro(1)

            # 1) Confirmação inicial do formulário
            safe_click((By.XPATH, "/html/body/div[2]/form/div[1]/div/div[1]/button[2]"))
            pausa_cadastro(2)
    
            # 2) Tenta fechar popup de duplicidade (se existir) — espera curta
            # (4s) pra nao segurar 20s em todo lead em que o popup nao aparece
            try:
                botao_dup = WebDriverWait(driver, 4).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="popVerificaDuplicidade"]/div/div/div[3]/button'))
                )
                botao_dup.click()
                pausa_cadastro(0.5)
                print(f"Lead duplicado encontrado para telefone {telefone}. Pulando {nome}.")
                resultados_cadastro.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'duplicado',
                    'Detalhe': 'Popup de duplicidade apareceu no cadastro.',
                })
                salvar_estado(index)
                break
            except Exception:
                pass
    
            # 3) Salvar
            safe_click((By.XPATH, '//*[@id="cmdSalva"]'))
            pausa_cadastro(3)

            # Confirma que o lead realmente entrou (busca o telefone no Fac);
            # so marca 'cadastrado' com confirmacao — sem ela vira erro p/ revisao.
            cadastro_confirmado, detalhe_confirmacao = telefone_existe_no_sigavi(telefone)
            if cadastro_confirmado:
                print(f"[CADASTRADO] {telefone}")
                resultados_cadastro.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'cadastrado',
                    'Detalhe': 'Lead cadastrado e confirmado no Sigavi.',
                })
            else:
                print(f"[ERRO] Salvar nao confirmou o cadastro de {nome} - {telefone}. Marcando para revisao.")
                resultados_cadastro.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'erro_cadastro',
                    'Detalhe': f'Cadastro nao confirmado apos salvar: {detalhe_confirmacao}',
                })
            salvar_estado(index)
            break
    
          except TimeoutException as e:
            # Timeout de elemento NAO e browser caido (TimeoutException herda de
            # WebDriverException) — sem este handler, qualquer elemento sumido
            # da pagina derrubava e recriava o Edge a toa, mascarando o erro real
            # como "Browser caiu" com Message vazio.
            if tentativa_lead == 1:
                print(f"[ERRO] Elemento nao apareceu no cadastro de {nome}: {e.msg or 'timeout'}. Tentando o lead de novo (2/2)...")
                continue
            print(f"[ERRO] Timeout de novo no lead {nome} na 2a tentativa. Pulando pro proximo.")
            resultados_cadastro.append({
                'Linha': index + 1,
                'Nome': nome,
                'Email': email_raw,
                'Telefone': telefone,
                'Status': 'erro_cadastro',
                'Detalhe': f"Elemento nao encontrado na pagina (timeout apos 2 tentativas): {e.msg or 'timeout'}",
            })
            salvar_estado(index)
          except (InvalidSessionIdException, WebDriverException) as e:
            print(f"[ERRO] Browser caiu — reconectando...")
            for tentativa_reconexao in range(3):
                try:
                    reconectar_browser()
                    break
                except Exception as re_err:
                    print(f"[ERRO] Reconexão falhou ({tentativa_reconexao+1}/3): {re_err}")
                    time.sleep(5)
            else:
                print("[ERRO] Não foi possível reconectar. Encerrando.")
                resultados_cadastro.append({
                    'Linha': index + 1,
                    'Nome': nome,
                    'Email': email_raw,
                    'Telefone': telefone,
                    'Status': 'erro_cadastro',
                    'Detalhe': f'Falha ao reconectar navegador: {e}',
                })
                salvar_estado(index, anunciar=True)
                raise SystemExit(1)
            if tentativa_lead == 1:
                print(f"Reconectado. Tentando o lead {nome} de novo (2/2)...")
                continue
            print(f"[ERRO] Browser caiu de novo no lead {nome}. Pulando pro proximo.")
            resultados_cadastro.append({
                'Linha': index + 1,
                'Nome': nome,
                'Email': email_raw,
                'Telefone': telefone,
                'Status': 'erro_cadastro',
                'Detalhe': f'Browser caiu durante cadastro (2 tentativas): {e}',
            })
            salvar_estado(index)
          except Exception as e:
            if tentativa_lead == 1:
                print(f"[ERRO] Falha ao cadastrar {nome}: {e}. Tentando o lead de novo (2/2)...")
                continue
            print(f"[ERRO] Falha ao cadastrar {nome} na 2a tentativa: {e}. Pulando pro proximo.")
            resultados_cadastro.append({
                'Linha': index + 1,
                'Nome': nome,
                'Email': email_raw,
                'Telefone': telefone,
                'Status': 'erro_cadastro',
                'Detalhe': f'{e} (apos 2 tentativas)',
            })
            salvar_estado(index)
    
        salvar_progresso(index, resultados_email, resultados_cadastro)
        if len(resultados_cadastro) % AUTOSAVE_EVERY == 0:
            _salvar_excel_resultado()

except KeyboardInterrupt:
    print("\n\nPausado pelo usuário. Salvando resultados...")
    ultimo_seguro = locals().get('index', _ultimo_index)
    salvar_estado(ultimo_seguro, anunciar=True)
    emitir_progresso()
    driver.quit()
    raise SystemExit(0)

print("Processamento concluído!")
driver.quit()
salvar_estado(len(df) - 1, anunciar=True)
emitir_progresso()

