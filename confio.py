import argparse
import re
import time
import json
import os
import unicodedata
import pandas as pd
import requests
from dotenv import load_dotenv

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
ULTIMO_ERRO_CONSULTA = None

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
if 'FONE2' in df.columns:
    df['FONE2'] = df['FONE2'].apply(
        lambda x: str(int(x)) if pd.notna(x) and str(x).strip() not in ('', 'nan') else ''
    )

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
    options = webdriver.EdgeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    if HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
    d = webdriver.Edge(options=options)
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
    # respostas curtas (< 5000 chars) são "sem resultado" — respostas longas têm JS do Kendo
    # que contém palavras como "não contém", "não é igual" que disparariam falso positivo
    if len(html) < 5000:
        sem_kw = ['não retornou', 'nao retornou', 'sem resultado', 'nenhum registro',
                  'nao foram encontrados', 'não foram encontrados']
        if any(k in html.lower() for k in sem_kw):
            return None

    # 1) células marcadas como dados pessoais (telefone preenchido)
    dados_pessoais = re.findall(r'<td[^>]*data-dados-pessoais="true"[^>]*>([^<]*)</td>', html)
    for celula in dados_pessoais:
        tel = _extrair_fone_de_texto(celula)
        if tel:
            return tel

    # 2) todas as <td> do tbody
    tbody = re.search(r'<tbody>(.*?)</tbody>', html, re.DOTALL)
    if tbody:
        for celula in re.findall(r'<td[^>]*>([^<]*)</td>', tbody.group(1)):
            tel = _extrair_fone_de_texto(celula)
            if tel:
                return tel

    return None

def buscar_telefone_por_email(session, csrf_token, email):
    global ULTIMO_ERRO_CONSULTA, req_csrf_token
    ULTIMO_ERRO_CONSULTA = None
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
    try:
        resp = session.post(
            f"{BASE_URL}/CRM/Fac/Busca",
            data=payload,
            headers={'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'Accept': '*/*'},
            timeout=30,
        )
        if resp.status_code == 200:
            # sessão expirada → servidor retorna página de login
            if 'ReturnUrl' in resp.url or 'Login' in resp.url or 'login' in resp.text[:500].lower():
                print(f"  [!] sessão expirada — renovando cookies...")
                for cookie in driver.get_cookies():
                    session.cookies.set(cookie['name'], cookie['value'])
                novo_token = obter_csrf_token(session)
                if novo_token:
                    req_csrf_token = novo_token
                else:
                    ULTIMO_ERRO_CONSULTA = 'Sessao expirada e CSRF token nao foi renovado.'
                resp = session.post(
                    f"{BASE_URL}/CRM/Fac/Busca",
                    data={**payload, '__RequestVerificationToken': req_csrf_token},
                    headers={'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'Accept': '*/*'},
                    timeout=30,
                )
                if 'ReturnUrl' in resp.url or 'Login' in resp.url or 'login' in resp.text[:500].lower():
                    ULTIMO_ERRO_CONSULTA = 'Sessao expirada mesmo apos renovar cookies.'
                    print(f"  [!] sessao ainda expirada apos renovar cookies")
                    return None
            tel = extrair_telefone_do_html(resp.text)
            if tel:
                print(f"  [✓] {tel}")
            else:
                print(f"  [✗] não encontrado ({len(resp.text)} chars)")
            return tel
        print(f"  [✗] status {resp.status_code}")
        ULTIMO_ERRO_CONSULTA = f'Status HTTP inesperado: {resp.status_code}'
    except Exception as e:
        print(f"  → Erro ao buscar email {email}: {e}")
        ULTIMO_ERRO_CONSULTA = str(e)
    return None

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
time.sleep(2)

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

# pré-calcula quantas linhas não têm telefone (para o contador)
if _email_col:
    if 'FONE2' in df.columns:
        _sem_fone = df['FONE2'].isna() | (df['FONE2'].astype(str).str.replace(r'\D', '', regex=True).str.len() < 11)
    else:
        _sem_fone = pd.Series([True] * len(df))
    _total_emails = int((df[_email_col].notna() & _sem_fone).sum())
else:
    _total_emails = 0
_email_count = 0

if MODE == 'consulta':
    print("Modo selecionado: somente consulta.")
else:
    print("Modo selecionado: somente cadastro.")
    driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
    time.sleep(2)

# mapa para tolerar variações de caixa no dicionário
mapa_corretores = {str(k).upper(): v for k, v in corretores_gerentes.items()}

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

        relatorios = {
            'cadastrados': df_todos[df_todos['Status'] == 'cadastrado'].reset_index(drop=True),
            'duplicados': df_todos[df_todos['Status'] == 'duplicado'].reset_index(drop=True),
            'nao_cadastrados': df_todos[df_todos['Status'] == 'nao_cadastrado'].reset_index(drop=True),
            'erros_cadastro': df_todos[df_todos['Status'] == 'erro_cadastro'].reset_index(drop=True),
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
                f"{len(relatorios['erros_cadastro'])} erro(s)."
            )


def salvar_estado(ultimo_index, anunciar=False):
    salvar_progresso(ultimo_index, resultados_email, resultados_cadastro)
    _salvar_excel_resultado(anunciar=anunciar)


if MODE == 'consulta':
    ultimo_processado = _ultimo_index
    try:
        for index, row in df.iterrows():
            if index <= _ultimo_index:
                continue

            if parada_solicitada():
                print("\nParada solicitada pela interface. Salvando resultados...")
                salvar_estado(ultimo_processado, anunciar=True)
                driver.quit()
                raise SystemExit(0)

            nome = str(row.get('NOME') or '').strip()
            email_raw = ''
            if _email_col:
                email_raw = str(row.get(_email_col) or '').strip()

            telefone_raw = str(row.get('FONE2') or '')
            telefone = re.sub(r'\D', '', telefone_raw)

            detalhe = ''
            status = 'nao_encontrado'

            if not email_raw:
                status = 'erro_consulta'
                detalhe = 'Email ausente na planilha.'
                print(f"[Linha {index + 1}] email ausente")
            elif len(telefone) >= 11:
                status = 'encontrado'
                detalhe = 'Telefone ja estava na planilha.'
                print(f"[Linha {index + 1}] {email_raw}")
                print(f"  [✓] {telefone} (planilha)")
            elif req_session and req_csrf_token:
                _email_count += 1
                print(f"[Email {_email_count}/{_total_emails}] {email_raw}")
                telefone = buscar_telefone_por_email(req_session, req_csrf_token, email_raw) or ''
                if len(telefone) >= 10:
                    status = 'encontrado'
                    detalhe = 'Telefone encontrado por email.'
                elif ULTIMO_ERRO_CONSULTA:
                    status = 'erro_consulta'
                    detalhe = ULTIMO_ERRO_CONSULTA
                else:
                    status = 'nao_encontrado'
                    detalhe = 'Telefone nao encontrado por email.'
            else:
                status = 'erro_consulta'
                detalhe = 'Sessao de consulta indisponivel.'

            resultados_email.append({
                'Linha': index + 1,
                'Nome': nome,
                'Email': email_raw,
                'Telefone': telefone if status == 'encontrado' else '',
                'Status': status,
                'Detalhe': detalhe,
            })

            ultimo_processado = index
            salvar_progresso(ultimo_processado, resultados_email, resultados_cadastro)
            if len(resultados_email) % AUTOSAVE_EVERY == 0:
                _salvar_excel_resultado()

        print("Consulta concluida!")
        salvar_estado(ultimo_processado, anunciar=True)
        driver.quit()
        raise SystemExit(0)

    except KeyboardInterrupt:
        print("\n\nPausado pelo usuario. Salvando resultados...")
        salvar_estado(ultimo_processado, anunciar=True)
        driver.quit()
        raise SystemExit(0)

try:
    for index, row in df.iterrows():
        if index <= _ultimo_index:
            continue
        if parada_solicitada():
            print("\nParada solicitada pela interface. Salvando resultados...")
            salvar_estado(index - 1, anunciar=True)
            driver.quit()
            raise SystemExit(0)

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

        gerente_planilha = str(row.get('GERENTE') or '').strip()
        if gerente_planilha:
            gerente  = gerente_planilha
            corretor = corretor_original_raw.strip() or "Corretor Inativo"
        elif corretor_original_norm in mapa_corretores:
            gerente  = mapa_corretores[corretor_original_norm]
            corretor = corretor_original_raw.strip()
        elif corretor_original_norm in mapa_corretores_base:
            gerente  = mapa_corretores_base[corretor_original_norm]
            corretor = corretor_original_raw.strip()
        else:
            gerente  = "Tabatanascimento"
            corretor = "Corretor Inativo"
    
        canal_setas = 4 if tipo_plantao_norm in canal_plantao_tipos else 1
        if tipo_plantao_norm in canal_carteira_tipos:
            canal_setas = 1
    
    
        try:
            # Página de busca (apenas preenche telefone; não valida duplicidade aqui)
            telefone_busca_locator = (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[1]/form/div[2]/div/div/div/div[10]/div[4]/input')
            telefone_elem_busca = None
            for tentativa in range(3):
                if parada_solicitada():
                    raise KeyboardInterrupt
                driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
                time.sleep(3)
                try:
                    telefone_elem_busca = wait_visible(telefone_busca_locator, timeout=20)
                    break
                except TimeoutException:
                    print(f"Página /CRM/Fac não carregou (tentativa {tentativa+1}/3). Tentando novamente...")
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
                continue
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
            time.sleep(2.5)  # aguarda carregamento da grade de resultados
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
                continue
    
            # navegação leve (como no seu script)
            ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
            time.sleep(0.5)
            ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
            time.sleep(0.5)
    
            # Vai direto ao cadastro
            driver.get('https://abyara.sigavi360.com.br/CRM/Fac/Cadastro')
            time.sleep(2)
    
            # === BLOCO CADASTRO ===
            wait_visible((By.ID, 'Nome')).send_keys(nome)
            time.sleep(2.5)
    
            # Abre o bloco de telefones
            safe_click((By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/div/a'))
            time.sleep(1)
    
            # Seleciona "Celular" no tipo (setas + enter)
            celular_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[1]/span[1]/span/span[1]')
            safe_click(celular_combo_locator)
            ActionChains(driver).send_keys(Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ENTER).perform()
    
            # Preenche número
            telefone_grid_input_locator = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[3]/input')
            tel_input = wait_visible(telefone_grid_input_locator)
            tel_input.click()
            tel_input.send_keys(telefone)
            time.sleep(1)
    
            # Adiciona telefone (ícone de +/confirmar)
            safe_click((By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[4]/a[1]/span'))
            time.sleep(1)
    
            # Canal (SMS)
            sms_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[1]/span[2]/span/span[1]')
            safe_click(sms_combo_locator)
            ac_canal = ActionChains(driver)
            for _ in range(canal_setas):
                ac_canal.send_keys(Keys.ARROW_DOWN)
            ac_canal.send_keys(Keys.ENTER).perform()
            time.sleep(1)
    
            # Mídia
            midia_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[2]/span[2]/span/span[1]')
            safe_click(midia_combo_locator)
            # seleciona conforme TIPO PLANTAO
            selecionar_midia_por_posicao(posicao_midia)
            time.sleep(1)
    
            # Equipe (gerente)
            equipe_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[1]/span[1]/span/span[1]')
            safe_click(equipe_combo_locator)
            ActionChains(driver).send_keys(gerente, Keys.ENTER).perform()
            time.sleep(0.8)
    
            # Corretor
            corretor_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/span[1]/span/span[1]')
            safe_click(corretor_combo_locator)
            ActionChains(driver).send_keys(corretor, Keys.ENTER).perform()
            time.sleep(0.8)
    
            # Abre modal "Imóvel/Origem"
            safe_click((By.XPATH, "/html/body/div[2]/form/div[3]/div/div/div[1]/div[2]/div[2]/a/span"))
            modal_container = wait_visible((By.XPATH, "/html/body/div[2]/div[2]/div/div"))
    
            # Seleciona a opção dentro do modal
            safe_click((By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[1]/div/div/label[2]"))
    
            # Preenche o código/descrição
            your_code_input_locator = (By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/input")
            your_code = wait_visible(your_code_input_locator)
            your_code.clear()
            your_code.send_keys("lume")
            time.sleep(0.5)
    
            # Confirma modal
            safe_click((By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/button"))
            time.sleep(1)
    
            # Botão que às vezes fica atrás de overlay
            safe_click((By.CSS_SELECTOR, "#dvImovelOrigemComando a"))
            time.sleep(1)
    
            # 1) Confirmação inicial do formulário
            safe_click((By.XPATH, "/html/body/div[2]/form/div[1]/div/div[1]/button[2]"))
            time.sleep(2)
    
            # 2) Tenta fechar popup de duplicidade (se existir)
            try:
                safe_click((By.XPATH, '//*[@id="popVerificaDuplicidade"]/div/div/div[3]/button'))
                time.sleep(0.5)
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
                continue
            except Exception:
                pass
    
            # 3) Salvar
            safe_click((By.XPATH, '//*[@id="cmdSalva"]'))
            time.sleep(3)
            print(f"[CADASTRADO] {telefone}")
            resultados_cadastro.append({
                'Linha': index + 1,
                'Nome': nome,
                'Email': email_raw,
                'Telefone': telefone,
                'Status': 'cadastrado',
                'Detalhe': 'Lead cadastrado no Sigavi.',
            })
    
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
            resultados_cadastro.append({
                'Linha': index + 1,
                'Nome': nome,
                'Email': email_raw,
                'Telefone': telefone,
                'Status': 'erro_cadastro',
                'Detalhe': f'Browser caiu durante cadastro: {e}',
            })
            salvar_estado(index)
            continue
        except Exception as e:
            print(f"[ERRO] Falha ao cadastrar {nome}: {e}")
            resultados_cadastro.append({
                'Linha': index + 1,
                'Nome': nome,
                'Email': email_raw,
                'Telefone': telefone,
                'Status': 'erro_cadastro',
                'Detalhe': str(e),
            })
            salvar_estado(index)
            continue
    
        salvar_progresso(index, resultados_email, resultados_cadastro)
        if len(resultados_cadastro) % AUTOSAVE_EVERY == 0:
            _salvar_excel_resultado()

except KeyboardInterrupt:
    print("\n\nPausado pelo usuário. Salvando resultados...")
    ultimo_seguro = locals().get('index', _ultimo_index)
    salvar_estado(ultimo_seguro, anunciar=True)
    driver.quit()
    raise SystemExit(0)

print("Processamento concluído!")
driver.quit()
salvar_estado(len(df) - 1, anunciar=True)

