# -*- coding: utf-8 -*-
"""Cliente da API REST do Sigavi — o que substitui o navegador (Selenium) no importador.

Por que curl e não requests/urllib: o IIS do Sigavi devolve HTTP 500 em QUALQUER rota
autenticada quando a requisição leva o header Accept-Encoding (urllib e fetch mandam
sozinhos). Parece queda geral, mas é o transporte. curl não manda — então é curl.

Regras que valem pra qualquer coisa que fale com esta API (medidas, não palpite):
  * Fac/Busca sem `StatusGestao:[1,2,3]` só devolve ficha ATIVA — com ele, acha finalizada.
  * Fac/Busca filtra de verdade por Numero, Telefone (match exato, sem o 55) e Email.
    `Cliente` é IGNORADO (devolve ficha de gente aleatória) — busca por nome não existe aqui.
  * A API não aguenta paralelismo: 2 em paralelo no máximo, senão vem 503.
  * `fac/salva` cria ficha NOVA a cada chamada e às vezes dá 500 depois de ter criado.
    Por isso NUNCA é retentado: quem chama confere pela busca antes de desistir.
"""
import json
import subprocess
import threading
import time
import unicodedata
import re

BASE_URL = "https://abyara.sigavi360.com.br"
STATUS_TODOS = [1, 2, 3]   # Em Atendimento + Atendimento Finalizado + o que estiver no meio


def normalizar(texto) -> str:
    """Sem acento, maiúsculo, só letra/número/espaço — pra casar nomes vindos de planilha."""
    if texto is None:
        return ''
    nfkd = unicodedata.normalize('NFKD', str(texto))
    sem = ''.join(c for c in nfkd if not unicodedata.combining(c)).upper()
    return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9]+', ' ', sem)).strip()


class ErroSigavi(Exception):
    pass


class SigaviAPI:
    def __init__(self, login, senha, login_reserva=None, senha_reserva=None, log=print):
        self._contas = [(login, senha)]
        if login_reserva and senha_reserva:
            self._contas.append((login_reserva, senha_reserva))
        self._conta = 0
        self._token = None
        self._lock = threading.Lock()
        self._log = log
        self._cache = {}

    # ── transporte ────────────────────────────────────────────────────────────
    def _curl(self, args, timeout):
        r = subprocess.run(["curl", "-s", "--max-time", str(timeout), "-w", "\n|%{http_code}", *args],
                           capture_output=True, text=True)
        corpo, _, code = r.stdout.rpartition("|")
        try:
            dados = json.loads(corpo.strip()) if corpo.strip() else None
        except ValueError:
            dados = corpo.strip()[:300]
        return int(code.strip() or 0), dados

    def autenticar(self):
        """Pega token com a conta atual; se a principal recusar, tenta a reserva (uma vez)."""
        with self._lock:
            while True:
                login, senha = self._contas[self._conta]
                for tentativa in range(3):
                    code, dados = self._curl([
                        "-X", "POST", f"{BASE_URL}/Sigavi/api/Acesso/Token",
                        "-H", "Content-Type: application/x-www-form-urlencoded",
                        "--data-urlencode", f"username={login}", "--data-urlencode", f"password={senha}",
                        "--data-urlencode", "grant_type=password"], 30)
                    if code == 200 and isinstance(dados, dict) and dados.get("access_token"):
                        self._token = dados["access_token"]
                        return self._token
                    if code in (400, 401):
                        break          # senha errada não melhora tentando de novo
                    time.sleep(5 * (tentativa + 1))
                if self._conta + 1 < len(self._contas):
                    self._conta += 1
                    self._log("Conta principal nao autenticou. Trocando pra conta reserva.")
                    continue
                raise ErroSigavi("Login no Sigavi recusado (usuario/senha) ou Sigavi fora do ar.")

    def _chamar(self, metodo, rota, corpo=None, timeout=60, idempotente=True):
        """(status, json). Leitura retenta 500/503/timeout com espera crescente; escrita não."""
        espera = 2
        tentativas = 6 if idempotente else 1
        ultimo = (0, None)
        for _ in range(tentativas):
            tok = self._token or self.autenticar()
            args = ["-X", metodo, f"{BASE_URL}{rota}", "-H", f"Authorization: bearer {tok}"]
            if corpo is not None:
                args += ["-H", "Content-Type: application/json", "-d", json.dumps(corpo)]
            code, dados = self._curl(args, timeout)
            ultimo = (code, dados)
            if code == 401:
                self._token = None
                continue
            if code == 200 or not idempotente:
                return code, dados
            time.sleep(espera)      # 500 intermitente do IIS / 503 app pool / timeout (0)
            espera = min(espera * 2, 30)
        return ultimo

    # ── leitura ───────────────────────────────────────────────────────────────
    def busca_facs(self, **filtros):
        """Lista de FACs (todas as situações). Numero/Telefone/Email. Levanta ErroSigavi se não respondeu."""
        corpo = {"StatusGestao": STATUS_TODOS, "Pagina": 0, "PaginaTamanho": 50, **filtros}
        code, dados = self._chamar("POST", "/Sigavi/api/Fac/Busca", corpo)
        if code != 200 or not isinstance(dados, list):
            raise ErroSigavi(f"Busca no Sigavi falhou (HTTP {code}).")
        return dados

    def midias(self):
        if "midias" not in self._cache:
            code, dados = self._chamar("GET", "/api/crm/fac/Midia")
            if code != 200 or not isinstance(dados, list):
                raise ErroSigavi(f"Nao consegui a lista de midias do Sigavi (HTTP {code}).")
            self._cache["midias"] = [{"Id": m["Id"], "Nome": (m.get("Nome") or "").strip()} for m in dados]
        return self._cache["midias"]

    def empreendimentos(self):
        """Todos os empreendimentos (a rota ignora filtro de nome — filtra-se aqui)."""
        if "empreendimentos" not in self._cache:
            code, dados = self._chamar("POST", "/api/produto/empreendimento/busca",
                                       {"Pagina": 0, "PaginaTamanho": 2000}, timeout=120)
            if code != 200 or not isinstance(dados, list):
                raise ErroSigavi(f"Nao consegui a lista de empreendimentos do Sigavi (HTTP {code}).")
            self._cache["empreendimentos"] = [
                {"Id": e.get("Id"), "Nome": (e.get("Nome") or e.get("Descricao") or "").strip()}
                for e in dados if e.get("Id")]
        return self._cache["empreendimentos"]

    def corretores(self, nome):
        """Autônomos cujo nome/apelido contém `nome` (o apelido do corretores.json é o NomeComercial)."""
        chave = ("corretor", normalizar(nome))
        if chave not in self._cache:
            code, dados = self._chamar("POST", "/api/Autonomo/Busca", {"Nome": nome, "Pagina": 0, "PaginaTamanho": 100})
            if code != 200 or not isinstance(dados, list):
                raise ErroSigavi(f"Busca de corretor no Sigavi falhou (HTTP {code}).")
            self._cache[chave] = [{"Id": c.get("Id"), "Nome": c.get("Nome") or "", "Apelido": c.get("NomeComercial") or "",
                                   "Equipe": c.get("Equipe") or "", "Ativo": bool(c.get("Ativo"))} for c in dados]
        return self._cache[chave]

    # ── escrita ───────────────────────────────────────────────────────────────
    def fac_salva(self, payload):
        """Cria UMA FAC. Nunca retenta. {ok, status, numero, erro}."""
        code, dados = self._chamar("POST", "/api/crm/fac/salva", payload, timeout=90, idempotente=False)
        erros = [e for e in (dados.get("Erros") or []) if e] if isinstance(dados, dict) else []
        recusado = isinstance(dados, dict) and (dados.get("Sucesso") is False or erros)
        numero = None
        if isinstance(dados, dict):
            for k in ("Numero", "IdFac", "Id"):
                try:
                    n = int(dados.get(k) or 0)
                except (TypeError, ValueError):
                    n = 0
                if n > 0:
                    numero = n
                    break
        ok = code == 200 and not recusado
        erro = None if ok else ("; ".join(erros) or (f"HTTP {code}" if code != 200 else "Sucesso:false"))
        return {"ok": ok, "status": code, "numero": numero, "erro": erro}
