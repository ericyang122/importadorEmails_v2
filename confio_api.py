# -*- coding: utf-8 -*-
"""Importador Sigavi — versão API (sem navegador). Substitui o confio.py (Selenium).

Mesmo contrato do confio.py, pra o app.py/tela/WhatsApp não perceberem a troca:
mesmos argumentos, mesmas linhas PROGRESS=/log que o app.js lê, mesmo progresso.json
(o "reprocessar erros" volta pela coluna Linha = índice + 1) e as MESMAS planilhas de
resultado (nomes, abas e colunas).

O que muda por baixo (29/09/2026):
  * login e buscas pela API REST (sigavi_api.py) — sem Chrome, sem cookie, sem XPath;
  * cadastro pelo `fac/salva` — nome, telefone, canal, mídia, corretor e empreendimento
    por ID, e o Sigavi devolve o número da ficha;
  * o empreendimento NÃO é mais fixo ("arvo"): vem da coluna da planilha ou do que foi
    escolhido na tela (SIGAVI_EMPREENDIMENTO_TELA) — decisão do Erick, 29/09;
  * busca só por NOME não existe na API (o filtro `Cliente` é ignorado e devolve gente
    aleatória): linha que só tem nome sai "nao encontrado" dizendo isso, em vez de
    arriscar homônimo.

SIGAVI_DRY_RUN=true faz o cadastro inteiro (duplicidade, corretor, canal, mídia,
empreendimento) SEM criar a ficha — pra testar planilha nova sem sujar o Sigavi.
"""
import argparse
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pandas as pd
from dotenv import load_dotenv

from sigavi_api import ErroSigavi, SigaviAPI, normalizar

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass


def parse_args():
    p = argparse.ArgumentParser(description="Importa leads de uma planilha para o Sigavi (via API).")
    p.add_argument("--excel", default=os.getenv("SIGAVI_EXCEL", "./abertos/abertos_cora.xlsx"))
    p.add_argument("--headless", action="store_true", help="Ignorado (nao ha navegador); aceito por compatibilidade.")
    p.add_argument("--result-dir", default=os.getenv("SIGAVI_RESULT_DIR", "resultados"))
    p.add_argument("--mode", choices=("consulta", "cadastro", "verificar"), default=os.getenv("SIGAVI_MODE", "consulta"))
    p.add_argument("--stop-file", default=os.getenv("SIGAVI_STOP_FILE"))
    p.add_argument("--progress-file", default=os.getenv("SIGAVI_PROGRESS_FILE"))
    p.add_argument("--autosave-every", type=int, default=int(os.getenv("SIGAVI_AUTOSAVE_EVERY", "10")))
    return p.parse_args()


ARGS = parse_args()
load_dotenv()
SIGAVI_LOGIN = os.getenv("SIGAVI_LOGIN")
SIGAVI_SENHA = os.getenv("SIGAVI_SENHA")
RESULT_DIR = ARGS.result_dir
MODE = ARGS.mode
STOP_FILE = ARGS.stop_file
AUTOSAVE_EVERY = max(1, ARGS.autosave_every)
# a API do Sigavi derruba com paralelismo (503 a partir de ~3); 2 é o medido que aguenta
WORKERS = max(1, int(os.getenv("SIGAVI_API_WORKERS", "2")))
# Só o que veio da TELA (o app.py passa em SIGAVI_EMPREENDIMENTO_TELA). De propósito NÃO lê
# o SIGAVI_EMPREENDIMENTO do .env: lá ainda mora o "arvo" do robô antigo, e ele entraria
# calado em toda linha sem empreendimento — o contrário do combinado (29/09).
EMPREENDIMENTO_PADRAO = (os.getenv("SIGAVI_EMPREENDIMENTO_TELA") or "").strip()
DRY_RUN = os.getenv("SIGAVI_DRY_RUN", "").strip().lower() in ("1", "true", "sim", "yes")

# Canal de atendimento do cadastro. Não há rota de lista de canais na API; os IDs vêm
# das FACs reais (conferido 29/09): "Carteira" do robô = CARTEIRA CORRETOR (ok do Erick).
CANAIS = {"Carteira": 1004, "Plantão de Vendas": 5}
CANAL_NOME_SIGAVI = {"Carteira": "CARTEIRA CORRETOR", "Plantão de Vendas": "Plantão de Vendas"}

if not SIGAVI_LOGIN or not SIGAVI_SENHA:
    print("ERRO: informe SIGAVI_LOGIN e SIGAVI_SENHA no .env ou no ambiente da execucao.")
    raise SystemExit(1)


# =========================
# PROGRESSO (mesmo formato do confio.py — o reprocessar depende dele)
# =========================
def carregar_progresso():
    if os.path.exists(PROGRESSO_FILE):
        with open(PROGRESSO_FILE, 'r', encoding='utf-8') as f:
            dados = json.load(f)
        print(f"Retomando do índice {dados['ultimo_index'] + 1} (progresso salvo encontrado).")
        return dados.get('ultimo_index', -1), dados.get('resultados_email', []), dados.get('resultados_cadastro', [])
    return -1, [], []


def salvar_progresso(ultimo_index, resultados_email, resultados_cadastro):
    os.makedirs(os.path.dirname(PROGRESSO_FILE) or '.', exist_ok=True)
    with open(PROGRESSO_FILE, 'w', encoding='utf-8') as f:
        json.dump({'ultimo_index': ultimo_index, 'modo': MODE, 'resultados_email': resultados_email,
                   'resultados_cadastro': resultados_cadastro}, f, ensure_ascii=False)


# =========================
# CORRETORES (corretores.json: apelido do Sigavi -> equipe)
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
# PLANILHA (idêntico ao confio.py: cabeçalho com título em cima, renames, telefone)
# =========================
arquivo_excel = ARGS.excel
if not os.path.exists(arquivo_excel):
    print(f"ERRO: planilha nao encontrada: {arquivo_excel}")
    raise SystemExit(1)

excel_file = pd.ExcelFile(arquivo_excel)
nome_planilha = excel_file.sheet_names[0]
_RE_CABECALHO = re.compile(r"(telefone|celular|fone|whats|phone|e.?mail|nome|cliente|^\s*fac\s*$|corretor)", re.IGNORECASE)


def _achar_linha_cabecalho(bruto, max_linhas=15):
    if any(_RE_CABECALHO.search(str(c)) for c in bruto.columns if not str(c).startswith("Unnamed")):
        return None
    for i in range(min(max_linhas, len(bruto))):
        valores = [str(v) for v in bruto.iloc[i].tolist() if str(v).strip().lower() not in ("", "nan", "none")]
        if sum(1 for v in valores if _RE_CABECALHO.search(v)) >= 2:
            return i + 1
    return None


def _limpar_colunas(df):
    df = df.rename(columns=lambda c: str(c).strip())
    vazias = [c for c in df.columns if str(c).startswith("Unnamed") and df[c].isna().all()]
    return df.drop(columns=vazias)   # linha em branco NÃO sai: Linha = posição no dataframe


df = pd.read_excel(arquivo_excel, sheet_name=nome_planilha)
_linha_cab = _achar_linha_cabecalho(df)
if _linha_cab:
    df = pd.read_excel(arquivo_excel, sheet_name=nome_planilha, header=_linha_cab)
    print(f"Cabecalho encontrado na linha {_linha_cab + 1} (tinha titulo em cima).")
df = _limpar_colunas(df).reset_index(drop=True)
print(f"Arquivo carregado: {arquivo_excel} | Aba: '{nome_planilha}' | {len(df)} linhas")

df = df.rename(columns={
    'CORRETOR ORIGEM': 'CORRETOR DE ORIGEM', 'TELEFONE': 'FONE2', 'NOME COMPLETO': 'NOME',
    'nome_cliente': 'NOME', 'celular': 'FONE2', 'corretor': 'CORRETOR DE ORIGEM',
    'origem': 'TIPO PLANTAO', 'gerente': 'GERENTE',
})


def normalizar_telefone_planilha(valor):
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
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

_nome_excel = os.path.splitext(os.path.basename(arquivo_excel.lstrip('./')))[0]
PROGRESSO_FILE = ARGS.progress_file or os.path.join(RESULT_DIR, f'progresso_{_nome_excel}_{MODE}.json')


def parada_solicitada():
    return bool(STOP_FILE and os.path.exists(STOP_FILE))


def _resultado_file(sufixo):
    return os.path.join(RESULT_DIR, f'resultado_{_nome_excel}_{sufixo}.xlsx')


def texto(valor):
    """Célula -> texto limpo. Célula vazia vira '' (o confio.py deixava virar 'nan')."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return ''
    s = str(valor).strip()
    return '' if s.lower() in ('nan', 'none') else s


def numero_fac(valor):
    """FAC da planilha. Coluna numérica com vazio vira float (488368.0) — o confio.py
    colava o .0 e buscava 4883680."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return ''
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return re.sub(r'\D', '', str(valor))


# =========================
# COLUNAS
# =========================
_email_col = next((c for c in df.columns if re.match(r'e.?mail', c, re.IGNORECASE)), None)
print(f"Coluna de email detectada: '{_email_col}'" if _email_col else "AVISO: Nenhuma coluna de email encontrada na planilha.")
_fac_col = next((c for c in df.columns if re.fullmatch(r'\s*(fac|n[º°o]?\.?\s*fac|numero|n[º°o])\s*', str(c), re.IGNORECASE)), None)
if _fac_col:
    print(f"Coluna de FAC detectada: '{_fac_col}'")
_nome_col = ('NOME' if 'NOME' in df.columns else
             next((c for c in df.columns if re.search(r'(nome|cliente)', str(c), re.IGNORECASE)), None))
if _nome_col:
    print(f"Coluna de nome detectada: '{_nome_col}'")
_tel_col = ('FONE2' if 'FONE2' in df.columns else
            next((c for c in df.columns if re.search(r'(telefone|celular|fone|whats|phone)', str(c), re.IGNORECASE)), None))
if MODE == 'verificar' and _tel_col:
    print(f"Coluna de telefone detectada: '{_tel_col}'")
_emp_col = next((c for c in df.columns if re.search(r'empreend', str(c), re.IGNORECASE)), None)
if MODE == 'cadastro':
    if _emp_col:
        print(f"Coluna de empreendimento detectada: '{_emp_col}'"
              + (f" (linha vazia usa '{EMPREENDIMENTO_PADRAO}', escolhido na tela)" if EMPREENDIMENTO_PADRAO else ""))
    elif EMPREENDIMENTO_PADRAO:
        print(f"Empreendimento escolhido na tela: '{EMPREENDIMENTO_PADRAO}' (a planilha nao tem coluna de empreendimento).")
    else:
        print("ERRO: a planilha nao tem coluna de empreendimento e nenhum foi escolhido na tela.")
        raise SystemExit(1)

# =========================
# RESULTADOS (mesmos arquivos/colunas do confio.py)
# =========================
COLUNAS_VERIFICAR = [
    'Linha', 'Nome', 'Telefone', 'Email', 'Status', 'Buscado por', 'Qtd FACs',
    'FAC', 'Situacao', 'Cadastro', 'Atualizacao', 'Equipe', 'Corretor',
    'Cliente no Sigavi', 'Telefone no Sigavi', 'Email no Sigavi', 'Canal', 'Midia',
    'Empreendimento', 'Situacao Atendimento', 'Motivo Finalizacao', 'Outras FACs', 'Detalhe',
]
COLUNAS = ['Linha', 'Nome', 'Email', 'Telefone', 'Status', 'Detalhe']

_ultimo_index, resultados_email, resultados_cadastro = carregar_progresso()
resultados_corretor_inativo = []


def _salvar_excel_resultado(anunciar=False):
    os.makedirs(RESULT_DIR, exist_ok=True)
    if MODE == 'verificar':
        t = pd.DataFrame(resultados_email, columns=COLUNAS_VERIFICAR)
        rel = {s: t[t['Status'] == st].drop(columns=['Status']).reset_index(drop=True)
               for s, st in (('com_cadastro', 'encontrado'), ('sem_cadastro', 'nao_encontrado'), ('erros_consulta', 'erro_consulta'))}
    elif MODE == 'consulta':
        t = pd.DataFrame(resultados_email, columns=COLUNAS)
        rel = {s: t[t['Status'] == st].reset_index(drop=True)
               for s, st in (('encontrados', 'encontrado'), ('nao_encontrados', 'nao_encontrado'), ('erros_consulta', 'erro_consulta'))}
    else:
        t = pd.DataFrame(resultados_cadastro, columns=COLUNAS)
        rel = {s: t[t['Status'] == st].reset_index(drop=True)
               for s, st in (('cadastrados', 'cadastrado'), ('duplicados', 'duplicado'),
                             ('nao_cadastrados', 'nao_cadastrado'), ('erros_cadastro', 'erro_cadastro'))}
        rel['corretores_inativos'] = pd.DataFrame(resultados_corretor_inativo, columns=COLUNAS)
    arquivos = []
    for sufixo, dataframe in rel.items():
        arquivo = _resultado_file(sufixo)
        dataframe.to_excel(arquivo, sheet_name=sufixo[:31], index=False)
        arquivos.append(arquivo)
    if anunciar:
        for a in arquivos:
            print(f"RESULT_FILE={a}")
        if MODE == 'verificar':
            print(f"Resumo verificacao: {len(rel['com_cadastro'])} com cadastro, {len(rel['sem_cadastro'])} sem cadastro, "
                  f"{len(rel['erros_consulta'])} erro(s).")
        elif MODE == 'consulta':
            print(f"Resumo consulta: {len(rel['encontrados'])} encontrado(s), {len(rel['nao_encontrados'])} nao encontrado(s), "
                  f"{len(rel['erros_consulta'])} erro(s).")
        else:
            print(f"Resumo cadastro: {len(rel['cadastrados'])} cadastrado(s), {len(rel['duplicados'])} duplicado(s), "
                  f"{len(rel['nao_cadastrados'])} nao cadastrado(s), {len(rel['erros_cadastro'])} erro(s), "
                  f"{len(rel['corretores_inativos'])} com corretor inativo.")


def salvar_estado(ultimo_index, anunciar=False):
    salvar_progresso(ultimo_index, resultados_email, resultados_cadastro)
    _salvar_excel_resultado(anunciar=anunciar)


def emitir_progresso():
    if MODE in ('consulta', 'verificar'):
        r = resultados_email
        suc = sum(1 for x in r if x['Status'] == 'encontrado')
        pen = sum(1 for x in r if x['Status'] == 'nao_encontrado')
        err = sum(1 for x in r if x['Status'] == 'erro_consulta')
    else:
        r = resultados_cadastro
        suc = sum(1 for x in r if x['Status'] == 'cadastrado')
        pen = sum(1 for x in r if x['Status'] in ('duplicado', 'nao_cadastrado'))
        err = sum(1 for x in r if x['Status'] == 'erro_cadastro')
    print("PROGRESS=" + json.dumps({'processados': len(r), 'total': len(df), 'sucessos': suc,
                                    'pendentes': pen, 'erros': err}), flush=True)


# =========================
# LOGIN (API)
# =========================
api = SigaviAPI(SIGAVI_LOGIN, SIGAVI_SENHA,
                os.getenv("SIGAVI_LOGIN_FALLBACK"), os.getenv("SIGAVI_SENHA_FALLBACK"), log=print)
print("Autenticando no Sigavi (API)...")
try:
    api.autenticar()
except ErroSigavi as e:
    print(f"ERRO: {e} Verifique usuario e senha do Sigavi.")
    raise SystemExit(1)
print("Login no Sigavi confirmado.")
print("Session de busca por email pronta.")   # o app.js acompanha esta linha

# =========================
# APOIO: telefone, data, FAC -> colunas
# =========================
def _e_fone_valido(d):
    if len(d) not in (10, 11):
        return False
    if not 11 <= int(d[:2]) <= 99:
        return False
    return not (len(d) == 11 and d[2] != '9')


def _extrair_fone(texto_fone):
    d = re.sub(r'\D', '', texto_fone or '')
    if d.startswith('55') and len(d) >= 12:
        d = d[2:]
    if _e_fone_valido(d):
        return d
    blocos = re.findall(r'\d+', texto_fone or '')
    for i in range(len(blocos)):
        acum = ''
        for j in range(i, min(i + 5, len(blocos))):
            acum += blocos[j]
            if len(acum) > 11:
                break
            if len(acum) >= 10:
                d2 = acum[2:] if acum.startswith('55') and len(acum) >= 12 else acum
                if _e_fone_valido(d2):
                    return d2
    return None


def telefones_da_celula(bruto):
    """Todos os telefones de uma célula. Corretor às vezes anota dois:
    "(11)98765-4321 e 91234-5678" — juntos viram 20 dígitos e nenhum robô reconhecia
    (linha 104 do Relatório Cora Pinheiros, 23/09). O 2º sem DDD herda o DDD do 1º."""
    texto_cel = str(bruto or '')
    if len(re.sub(r'\D', '', texto_cel)) <= 13:
        return [texto_cel]                       # um número só (com ou sem 55): segue o caminho de sempre
    achados, ddd = [], ''
    for pedaco in re.split(r'\s*(?:/|;|,|\be\b|\bou\b|\|)\s*', texto_cel):
        d = re.sub(r'\D', '', pedaco)
        if d.startswith('55') and len(d) in (12, 13):
            d = d[2:]
        if len(d) in (8, 9) and ddd:
            d = ddd + d
        if len(d) in (10, 11):
            ddd = d[:2]
            achados.append(d)
    return achados


def _variantes_telefone(bruto):
    """Tira 55/zero de operadora; celular de 11 dígitos também sem o 9 (ficha antiga)."""
    d = re.sub(r'\D', '', str(bruto or ''))
    if d.startswith('55') and len(d) in (12, 13):
        d = d[2:]
    if d.startswith('0') and len(d) in (11, 12):
        d = d[1:]
    if len(d) not in (10, 11):
        return []
    v = [d]
    if len(d) == 11 and d[2] == '9':
        v.append(d[:2] + d[3:])
    return v


def _data_iso(valor):
    """'2026-09-21T10:25:00' -> datetime (pra ordenar) e '21/09/2026 10:25' (pra planilha)."""
    if not valor:
        return datetime.min, ''
    try:
        d = datetime.fromisoformat(str(valor)[:19])
        return d, d.strftime('%d/%m/%Y %H:%M')
    except ValueError:
        return datetime.min, str(valor)


def _fac_para_colunas(f):
    return {
        'FAC': str(f.get('Numero') or ''), 'Situacao': f.get('StatusGestao') or '',
        'Cadastro': _data_iso(f.get('Cadastro'))[1], 'Atualizacao': _data_iso(f.get('Atualizacao'))[1],
        'Equipe': f.get('Equipe') or '', 'Corretor': f.get('Corretor') or '',
        'Cliente no Sigavi': f.get('Cliente') or '', 'Telefone no Sigavi': f.get('Telefone') or '',
        'Email no Sigavi': f.get('Email') or '', 'Canal': f.get('Canal') or '', 'Midia': f.get('Midia') or '',
        'Empreendimento': f.get('Produto') or '', 'Situacao Atendimento': f.get('SituacaoAtendimento') or '',
        'Motivo Finalizacao': f.get('MotivoFinalizacao') or '',
    }


def _mais_recentes(facs):
    return sorted(facs, key=lambda f: _data_iso(f.get('Atualizacao'))[0], reverse=True)


SO_NOME = 'Busca so por nome nao existe na API do Sigavi (devolve homonimo); precisa de FAC, telefone ou email.'

# =========================
# CONSULTA: achar telefone (e email) — FAC > email
# =========================
def _processar_linha_consulta(index, row):
    nome = texto(row.get(_nome_col)) if _nome_col else texto(row.get('NOME'))
    email_raw = texto(row.get(_email_col)) if _email_col else ''
    fac = numero_fac(row.get(_fac_col)) if _fac_col else ''
    telefone = re.sub(r'\D', '', texto(row.get('FONE2')))

    def res(status, tel, email, detalhe):
        return {'Linha': index + 1, 'Nome': nome, 'Email': email, 'Telefone': tel, 'Status': status, 'Detalhe': detalhe}

    if len(telefone) >= 11 and email_raw:
        return res('encontrado', telefone, email_raw, 'Telefone e email ja estavam na planilha.')
    if fac:
        criterio, filtro = f'FAC {fac}', {'Numero': fac}
    elif email_raw:
        criterio, filtro = 'email', {'Email': email_raw}
    elif nome:
        return res('nao_encontrado', telefone, email_raw, SO_NOME)
    else:
        return res('erro_consulta', telefone, email_raw, 'Linha sem FAC, email nem nome para buscar.')

    try:
        facs = _mais_recentes(api.busca_facs(**filtro))
    except ErroSigavi as e:
        return res('erro_consulta', telefone, email_raw, str(e))
    tel_enc = next((t for t in (_extrair_fone(f.get('Telefone')) for f in facs) if t), None)
    email_enc = next((f.get('Email').strip() for f in facs if (f.get('Email') or '').strip()), '')
    tel_final = tel_enc or telefone
    email_final = email_raw or email_enc
    if len(tel_final) >= 10:
        return res('encontrado', tel_final, email_final, f'Encontrado por {criterio}.')
    if facs:
        return res('nao_encontrado', tel_final, email_final, f'Cadastro encontrado por {criterio}, mas sem telefone no Sigavi.')
    return res('nao_encontrado', tel_final, email_final, f'Nao encontrado por {criterio}.')


# =========================
# VERIFICAR: já tem ficha? — FAC > telefone(s) > email
# =========================
def _processar_linha_verificar(index, row):
    nome = texto(row.get(_nome_col)) if _nome_col else ''
    email_raw = texto(row.get(_email_col)) if _email_col else ''
    fac = numero_fac(row.get(_fac_col)) if _fac_col else ''
    tel_raw = normalizar_telefone_planilha(row.get(_tel_col)) if _tel_col else ''
    base = {c: '' for c in COLUNAS_VERIFICAR}
    base.update({'Linha': index + 1, 'Nome': nome, 'Telefone': tel_raw, 'Email': email_raw})

    tentativas = []
    if fac:
        tentativas.append((f'FAC {fac}', {'Numero': fac}))
    for tel in telefones_da_celula(tel_raw):
        for v in _variantes_telefone(tel):
            tentativas.append((f'telefone {v}', {'Telefone': v}))
    if email_raw:
        tentativas.append(('email', {'Email': email_raw}))
    if not tentativas:
        if nome:
            return {**base, 'Status': 'nao_encontrado', 'Qtd FACs': 0, 'Detalhe': SO_NOME}
        return {**base, 'Status': 'erro_consulta', 'Detalhe': 'Linha sem FAC, telefone, email nem nome.'}

    erros = []
    for criterio, filtro in tentativas:
        try:
            facs = api.busca_facs(**filtro)
        except ErroSigavi as e:
            erros.append(f'{criterio}: {e}')
            continue
        if facs:
            facs = _mais_recentes(facs)
            principal = _fac_para_colunas(facs[0])
            outras = ', '.join(f"{f.get('Numero')} ({f.get('StatusGestao') or ''}, {f.get('Corretor') or ''})" for f in facs[1:])
            detalhe = f'Encontrado por {criterio}.'
            if len(facs) >= 50:
                detalhe += ' O Sigavi tem 50 FACs ou mais; a lista mostra as 50 primeiras.'
            return {**base, **principal, 'Status': 'encontrado', 'Buscado por': criterio,
                    'Qtd FACs': len(facs), 'Outras FACs': outras, 'Detalhe': detalhe}
    if erros and len(erros) == len(tentativas):
        return {**base, 'Status': 'erro_consulta', 'Detalhe': ' | '.join(erros)[:300]}
    detalhe = f"Nenhuma FAC por {', '.join(c for c, _ in tentativas)}."
    if erros:
        detalhe += ' (falhou: ' + ' | '.join(erros)[:200] + ')'
    return {**base, 'Status': 'nao_encontrado', 'Qtd FACs': 0, 'Detalhe': detalhe}


if MODE in ('consulta', 'verificar'):
    print("Modo selecionado: verificar cadastro (so leitura)." if MODE == 'verificar' else "Modo selecionado: somente consulta.")
    processar_linha = _processar_linha_verificar if MODE == 'verificar' else _processar_linha_consulta
    ultimo_processado = _ultimo_index
    pendentes = [(i, r) for i, r in df.iterrows() if i > _ultimo_index]
    bloco_tam = max(AUTOSAVE_EVERY, WORKERS)
    total_pendentes = len(pendentes)
    processados = 0
    marcas = {'encontrado': '✓', 'nao_encontrado': '✗', 'erro_consulta': '!'}
    print(f"Consulta paralela: {total_pendentes} linha(s) a processar, {WORKERS} em paralelo.")
    emitir_progresso()
    try:
        for inicio in range(0, total_pendentes, bloco_tam):
            if parada_solicitada():
                print("\nParada solicitada pela interface. Salvando resultados...")
                salvar_estado(ultimo_processado, anunciar=True)
                emitir_progresso()
                raise SystemExit(0)
            bloco = pendentes[inicio:inicio + bloco_tam]
            feitos = {}
            with ThreadPoolExecutor(max_workers=WORKERS) as ex:
                futuros = {ex.submit(processar_linha, i, r): i for i, r in bloco}
                for fut in as_completed(futuros):
                    feitos[futuros[fut]] = fut.result()
            for i, _r in bloco:
                resultado = feitos[i]
                resultados_email.append(resultado)
                processados += 1
                marca = marcas.get(resultado['Status'], '?')
                if MODE == 'verificar':
                    quem = resultado['Telefone'] or resultado['Email'] or resultado['Nome'] or '-'
                    print(f"[{processados}/{total_pendentes}] linha {resultado['Linha']} "
                          f"{str(quem).replace(' ', '_')} [{marca}] {('FAC_' + resultado['FAC']) if resultado['FAC'] else ''}")
                else:
                    print(f"[{processados}/{total_pendentes}] linha {resultado['Linha']} "
                          f"{resultado['Email']} [{marca}] {resultado['Telefone']}")
                ultimo_processado = i
            salvar_estado(ultimo_processado)
            emitir_progresso()
        print("Consulta concluida!")
        salvar_estado(ultimo_processado, anunciar=True)
        emitir_progresso()
        raise SystemExit(0)
    except KeyboardInterrupt:
        print("\n\nPausado pelo usuario. Salvando resultados...")
        salvar_estado(ultimo_processado, anunciar=True)
        emitir_progresso()
        raise SystemExit(0)

# =========================
# CADASTRO
# =========================
print("Modo selecionado: somente cadastro." + (" (SIMULACAO: nenhuma ficha sera criada)" if DRY_RUN else ""))
mapa_corretores = {str(k).upper(): v for k, v in corretores_gerentes.items()}
equipes_oficiais = sorted(set(corretores_gerentes.values()))
mapa_corretores_base = {}
for k, v in mapa_corretores.items():
    b = re.split(r'\s*[-|]\s*', k)[0].strip()
    if b and b not in mapa_corretores_base:
        mapa_corretores_base[b] = v

EQUIPE_INATIVO, CORRETOR_INATIVO = "Tabatanascimento", "Corretor Inativo"
MIDIA_AJUSTES = {normalizar("WHATS APP"): "WhatsApp", normalizar("INDICACAO FAMILIA AMIGO"): "Indicação",
                 normalizar("INDICACAO DO CORRETOR"): "Indicação"}
MIDIA_DEFAULT = "Outros"
_TIPOS_CARTEIRA = {normalizar(x) for x in ("INDICACAO", "INDICAÇÃO", "IND. CORRETOR", "IND CORRETOR", "CARTEIRA")}


def canal_por_tipo(tipo_norm):
    if tipo_norm and ('INDICACAO' in tipo_norm or 'CARTEIRA' in tipo_norm or tipo_norm in _TIPOS_CARTEIRA):
        return "Carteira"
    return "Plantão de Vendas"


def inativo(index, nome, email_raw, telefone, detalhe):
    resultados_corretor_inativo.append({'Linha': index + 1, 'Nome': nome, 'Email': email_raw, 'Telefone': telefone,
                                        'Status': 'corretor_inativo', 'Detalhe': detalhe})


def resolver_equipe_corretor(index, nome, email_raw, telefone, bruto, norm):
    """(equipe, apelido) pelo corretores.json — mesma regra do confio.py."""
    if norm in mapa_corretores:
        equipe, corretor = mapa_corretores[norm], bruto.strip()
    elif norm in mapa_corretores_base:
        equipe, corretor = mapa_corretores_base[norm], bruto.strip()
        print(f"Corretor '{bruto}' encontrado via nome base (sem sufixo). Gerente: {equipe}")
    else:
        print(f"Corretor '{bruto}' nao encontrado no corretores.json (inativo). Cadastrando como equipe {EQUIPE_INATIVO} / {CORRETOR_INATIVO}.")
        inativo(index, nome, email_raw, telefone,
                f"Corretor '{bruto.strip() or '(vazio)'}' nao encontrado no corretores.json; cadastrado como {EQUIPE_INATIVO}/{CORRETOR_INATIVO}.")
        return EQUIPE_INATIVO, CORRETOR_INATIVO
    en = normalizar(equipe)
    if en and all(normalizar(e) != en for e in equipes_oficiais):
        cands = [e for e in equipes_oficiais if en in normalizar(e) or normalizar(e) in en]
        if len(cands) == 1:
            print(f"Equipe '{equipe}' ajustada para '{cands[0]}' (nome oficial do corretores.json).")
            equipe = cands[0]
    return equipe, corretor


_ids_corretor = {}


def id_do_corretor(apelido, equipe):
    """Id no Sigavi pelo apelido (NomeComercial). Só corretor ATIVO; na dúvida, o da equipe."""
    chave = (normalizar(apelido), normalizar(equipe))
    if chave not in _ids_corretor:
        cands = [c for c in api.corretores(apelido)
                 if normalizar(apelido) in (normalizar(c['Apelido']), normalizar(c['Nome']))]
        ativos = [c for c in cands if c['Ativo']] or ([c for c in cands if c['Apelido'] == CORRETOR_INATIVO] if apelido == CORRETOR_INATIVO else [])
        da_equipe = [c for c in ativos if normalizar(c['Equipe']) == normalizar(equipe)]
        escolhido = (da_equipe or ativos or [None])[0]
        _ids_corretor[chave] = escolhido['Id'] if escolhido else None
    return _ids_corretor[chave]


def escolher_midia(midia_raw):
    lista = api.midias()
    por_norm = {}
    for m in lista:
        por_norm.setdefault(normalizar(m['Nome']), m)
    cands = [c for c in (MIDIA_AJUSTES.get(normalizar(midia_raw)), midia_raw) if c]
    for c in cands:
        if normalizar(c) in por_norm:
            return por_norm[normalizar(c)]
    for c in cands:
        cj = normalizar(c).replace(' ', '')
        for tn, m in por_norm.items():
            tj = tn.replace(' ', '')
            if cj and tj and (tj in cj or cj in tj):
                return m
    return por_norm.get(normalizar(MIDIA_DEFAULT))


_ids_emp = {}


def escolher_empreendimento(nome_emp):
    """(Id, Nome) do empreendimento pelo nome; erro com texto claro se não achar ou se for ambíguo."""
    n = normalizar(nome_emp)
    if n in _ids_emp:
        return _ids_emp[n]
    lista = api.empreendimentos()
    exatos = [e for e in lista if normalizar(e['Nome']) == n]
    if len(exatos) == 1:
        r = (exatos[0]['Id'], exatos[0]['Nome'], None)
    else:
        parecidos = [e for e in lista if n and (n in normalizar(e['Nome']) or normalizar(e['Nome']) in n)] if not exatos else exatos
        if len(parecidos) == 1:
            r = (parecidos[0]['Id'], parecidos[0]['Nome'], None)
        elif parecidos:
            r = (None, None, f"Empreendimento '{nome_emp}' bate com varios no Sigavi: "
                             + ', '.join(e['Nome'] for e in parecidos[:4]) + '. Escreva o nome completo.')
        else:
            r = (None, None, f"Empreendimento '{nome_emp}' nao encontrado no Sigavi.")
    _ids_emp[n] = r
    return r


def registrar(index, nome, email_raw, telefone, status, detalhe):
    resultados_cadastro.append({'Linha': index + 1, 'Nome': nome, 'Email': email_raw, 'Telefone': telefone,
                                'Status': status, 'Detalhe': detalhe})
    salvar_estado(index)


def cadastrar_linha(index, row):
    nome = texto(row.get('NOME'))
    email_raw = texto(row.get(_email_col)) if _email_col else ''
    fones = telefones_da_celula(texto(row.get('FONE2')))
    telefone = re.sub(r'\D', '', fones[0]) if fones else ''   # dois na célula: vale o primeiro
    if telefone.startswith('55') and len(telefone) in (12, 13):
        telefone = telefone[2:]          # o Sigavi guarda e busca SEM o 55
    if len(telefone) != 11:
        print(f"[NAO CADASTRADO] linha {index + 1}: telefone ausente ou invalido.")
        return registrar(index, nome, email_raw, telefone, 'nao_cadastrado', 'Telefone ausente ou invalido na planilha.')

    # empreendimento: coluna da linha > escolhido na tela
    nome_emp = (texto(row.get(_emp_col)) if _emp_col else '') or EMPREENDIMENTO_PADRAO
    if not nome_emp:
        return registrar(index, nome, email_raw, telefone, 'erro_cadastro',
                         'Linha sem empreendimento (coluna vazia e nenhum escolhido na tela).')
    id_emp, nome_emp_sigavi, erro_emp = escolher_empreendimento(nome_emp)
    if erro_emp:
        print(f"[ERRO] linha {index + 1}: {erro_emp}")
        return registrar(index, nome, email_raw, telefone, 'erro_cadastro', erro_emp)

    bruto = texto(row.get('CORRETOR DE ORIGEM'))
    equipe, apelido = resolver_equipe_corretor(index, nome, email_raw, telefone, bruto, re.sub(r'\s+', ' ', bruto).strip().upper())
    id_corretor = id_do_corretor(apelido, equipe)
    if not id_corretor and apelido != CORRETOR_INATIVO:
        # existe no corretores.json mas não está ativo no Sigavi (o caso JOTACE/Logan de 11/06)
        inativo(index, nome, email_raw, telefone,
                f"Corretor '{apelido}' nao esta ativo no Sigavi (equipe {equipe}); cadastrado como {EQUIPE_INATIVO}/{CORRETOR_INATIVO}.")
        equipe, apelido = EQUIPE_INATIVO, CORRETOR_INATIVO
        id_corretor = id_do_corretor(apelido, equipe)
    if not id_corretor:
        return registrar(index, nome, email_raw, telefone, 'erro_cadastro',
                         f"Corretor '{apelido}' (equipe {equipe}) nao foi encontrado no Sigavi.")

    tipo = texto(row.get('TIPO PLANTAO')) or texto(row.get('TIPO'))
    canal = canal_por_tipo(normalizar(tipo))
    print(f"Canal de Atendimento: pediu '{canal}' -> selecionou '{CANAL_NOME_SIGAVI[canal]}'")
    midia_raw = texto(row.get('MIDIA')) or texto(row.get('MÍDIA'))
    midia = escolher_midia(midia_raw)
    if not midia:
        return registrar(index, nome, email_raw, telefone, 'erro_cadastro', f"Midia '{midia_raw}' nao existe no Sigavi (nem '{MIDIA_DEFAULT}').")
    print(f"Midia: planilha '{midia_raw or '(vazio)'}' -> selecionou '{midia['Nome']}'")

    # duplicidade: qualquer FAC com esse telefone, em qualquer situação
    try:
        antes = api.busca_facs(Telefone=telefone)
    except ErroSigavi as e:
        return registrar(index, nome, email_raw, telefone, 'erro_cadastro', f'Nao consegui checar duplicidade: {e}')
    if antes:
        print(f"[DUPLICADO] {telefone}")
        return registrar(index, nome, email_raw, telefone, 'duplicado', 'Telefone ja encontrado no Sigavi antes do cadastro.')

    payload = {
        "IdCanal": CANAIS[canal], "IdMidia": midia['Id'], "IdCorretor": id_corretor, "IdProduto": id_emp,
        "Cliente": nome or "Sem nome", "Telefone": telefone, "Segmento": "Novos",
        "Mensagem": f"Importado da planilha {_nome_excel}",
    }
    if re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', email_raw):
        payload["Email"] = email_raw
    if DRY_RUN:
        print(f"[CADASTRADO] {telefone} (simulado)")
        return registrar(index, nome, email_raw, telefone, 'cadastrado',
                         f"SIMULACAO: criaria FAC com {apelido}/{equipe}, {CANAL_NOME_SIGAVI[canal]}, {midia['Nome']}, {nome_emp_sigavi}.")

    r = api.fac_salva(payload)
    # confirma pela busca (o número que volta pode ser o Id interno, não o da FAC)
    numero = None
    for espera in (1, 2, 4, 8):
        time.sleep(espera)
        try:
            depois = api.busca_facs(Telefone=telefone)
        except ErroSigavi:
            continue
        if depois:
            numero = _mais_recentes(depois)[0].get('Numero')
            break
    if numero:
        obs = '' if r['ok'] else f" (o Sigavi respondeu erro — {r['erro']} — mas a ficha foi criada)"
        print(f"[CADASTRADO] {telefone}")
        return registrar(index, nome, email_raw, telefone, 'cadastrado', f'Lead cadastrado e confirmado no Sigavi (FAC {numero}).{obs}')
    if r['ok']:
        print(f"[ERRO] Salvar nao confirmou: linha {index + 1}")
        return registrar(index, nome, email_raw, telefone, 'erro_cadastro',
                         'Cadastro nao confirmado apos salvar: Telefone nao apareceu na busca apos salvar')
    print(f"[ERRO] linha {index + 1}: Sigavi recusou o cadastro: {r['erro']}")
    return registrar(index, nome, email_raw, telefone, 'erro_cadastro', f"Sigavi recusou o cadastro: {r['erro']}")


try:
    for index, row in df.iterrows():
        if index <= _ultimo_index:
            continue
        if parada_solicitada():
            print("\nParada solicitada pela interface. Salvando resultados...")
            salvar_estado(index - 1, anunciar=True)
            emitir_progresso()
            raise SystemExit(0)
        emitir_progresso()
        try:
            cadastrar_linha(index, row)
        except ErroSigavi as e:
            print(f"[ERRO] linha {index + 1}: {e}")
            registrar(index, texto(row.get('NOME')), '', re.sub(r'\D', '', texto(row.get('FONE2'))), 'erro_cadastro', str(e))
    print("Processamento concluído!")
    salvar_estado(len(df) - 1, anunciar=True)
    emitir_progresso()
except KeyboardInterrupt:
    print("\n\nPausado pelo usuario. Salvando resultados...")
    salvar_estado(index, anunciar=True)
    emitir_progresso()
    raise SystemExit(0)
