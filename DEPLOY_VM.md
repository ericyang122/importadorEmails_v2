# DEPLOY NA VM — Importador Sigavi (guia técnico)

> **Pra quem é este arquivo:** pro Claude Code (ou dev) que vai configurar a VM da empresa.
> Objetivo: deixar o Importador Sigavi rodando 24/7, acessível por um link, com WhatsApp funcionando.
> **Atualizado em 18/06/2026** — plataforma definida como **Ubuntu Server** (antes a ideia era Win7, descartada).

---

## 0. ✅ PLATAFORMA DEFINIDA: Ubuntu Server 24.04 LTS

A VM será fornecida pelo infra (Eric Diniz), no servidor do trabalho, **ligada 24/7 (com nobreak)**.

**Specs pedidas ao infra:**
- **Ubuntu Server 24.04 LTS**
- **8 GB RAM** (pico de ~5 automações simultâneas na apresentação; cada automação ~700 MB + base ~2 GB)
- **4 vCPUs** | **40 GB SSD** | **Docker liberado**

> Histórico: a 1ª ideia era uma VM Windows 7, que NÃO suportava a stack (Python novo, Docker, Chrome). Ubuntu resolve tudo — é o cenário recomendado.

---

## 1. O QUE É O SISTEMA (visão geral)

RPA web que lê planilhas, consulta/cadastra leads no **Sigavi** (sistema web da Abyara) via navegador automatizado, gera planilhas de resultado e **notifica por WhatsApp**.

**Componentes que precisam estar de pé:**
1. **App Flask** (`app.py`) — interface web + orquestração. Servido por **waitress na porta 5000**.
2. **RPA Selenium** (`confio.py`) — controla o navegador (Chrome/Chromium **headless**) pra operar o Sigavi.
3. **Evolution API** (pasta `evolution/`) — envia o WhatsApp. Roda em **Docker**, porta **8080** (3 containers).
4. **Acesso** — o time acessa o link. Na VM da empresa, normalmente pela **rede interna** (`http://IP-DA-VM:5000`). Opcional: **ngrok** se precisar de acesso externo/HTTPS.

**Fluxo:** usuário acessa o link → login (senha `APP_PASSWORD`) → sobe planilha → o app dispara o Selenium no Sigavi (na própria VM, headless) → gera resultado → manda resumo/planilha no WhatsApp via Evolution.

---

## 2. PRÉ-REQUISITOS — instalar na VM Ubuntu

```bash
sudo apt update && sudo apt upgrade -y
```

### 2.1 Python + venv
```bash
sudo apt install -y python3 python3-venv python3-pip
# Ubuntu 24.04 já vem com Python 3.12 — compatível com as libs do projeto.
```

### 2.2 Git
```bash
sudo apt install -y git
```

### 2.3 Navegador pro Selenium — Google Chrome (recomendado)
> Mais confiável que o Chromium-snap do Ubuntu. O Selenium (versão do requirements) **baixa o driver sozinho** (Selenium Manager) — NÃO precisa instalar chromedriver à mão.
```bash
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo apt install -y ./google-chrome-stable_current_amd64.deb
google-chrome --version   # confirmar que instalou
```
⚠️ No `.env` setar **`SIGAVI_BROWSER=chrome`** (ver seção 4) — o código já suporta isso.

### 2.4 Docker + Docker Compose (pra Evolution API)
```bash
curl -fsSL https://get.docker.com | sudo sh
sudo usermod -aG docker $USER     # rodar docker sem sudo (relogar depois)
docker --version && docker compose version
```

### 2.5 ngrok (OPCIONAL — só se for expor pra fora da rede)
```bash
# se o acesso for só interno (IP da VM), PULA esta etapa
ngrok config add-authtoken <TOKEN_DA_CONTA_DO_ERICK>
```

---

## 3. INSTALAR O PROJETO (com venv)

```bash
git clone <repo> importadorEmails_v2
cd importadorEmails_v2
git checkout teste-automacao        # confirmar com o Erick a branch/commit certos

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# requirements: selenium, pandas, requests, openpyxl, python-dotenv, Flask, waitress
```

---

## 4. CONFIGURAR OS DOIS `.env` (segredos — NÃO estão no git)

> Há **dois** `.env`: um do app (raiz) e um da Evolution (`evolution/`). Copiar dos `.env.example` e preencher. **Pedir os valores reais ao Erick.**

### 4.1 `.env` (raiz) — config do app
```bash
cp .env.example .env
```
Campos que importam:
- **`SIGAVI_BROWSER=chrome`** ← ESSENCIAL no Ubuntu (usa o Chrome instalado em vez do Edge)
- `SIGAVI_HEADLESS=true` ← garante headless (a VM não tem tela)
- `APP_PASSWORD` — senha de acesso. **Trocar pra senha forte antes de expor.**
- `APP_SECRET_KEY` — chave aleatória grande (sessão Flask).
- `SESSION_COOKIE_SECURE` — `true` SÓ se acessar via HTTPS (ngrok). Na rede interna por HTTP, deixar `false`.
- `SIGAVI_LOGIN` / `SIGAVI_SENHA` — credenciais do Sigavi (a tela também pede).
- `EVOLUTION_API_URL` / `EVOLUTION_API_KEY` / `EVOLUTION_INSTANCE` / `WHATSAPP_DESTINO` — WhatsApp (ver 4.2).
- `SIGAVI_CADASTRO_DELAY` — fator de velocidade do cadastro (calibrar com leads reais).

### 4.2 `evolution/.env` — config da Evolution API
```bash
cp evolution/.env.example evolution/.env
```
- Preencher `AUTHENTICATION_API_KEY` (= `EVOLUTION_API_KEY` do app) e `POSTGRES_PASSWORD`.
- ⚠️ A senha do Postgres **só vale na 1ª criação do volume**. Se o volume já existir, manter a original.

---

## 5. SUBIR A EVOLUTION (WhatsApp) — Docker

```bash
cd evolution
docker compose up -d        # sobe evolution_api, evolution_postgres, evolution_redis
docker compose ps           # confirmar os 3 "Up"
curl http://localhost:8080  # deve responder "Welcome to the Evolution API"
```

### 5.1 ⚠️ PAREAR O WHATSAPP NA VM (passo manual obrigatório)
1. Acessar o painel `http://IP-DA-VM:8080/manager` (ou via túnel) — **a porta 8080 não deve ficar exposta pra internet sem proteção.**
2. Selecionar/criar a instância (nome em `EVOLUTION_INSTANCE`, ex.: `ericyang`).
3. **Escanear o QR code com o celular do Erick** (o número que envia as mensagens).
4. Confirmar conexão e mandar uma mensagem de teste pelo app.

> Envio é **silencioso**: se a Evolution cair, o job roda e gera planilhas, mas o resumo não chega no zap, sem erro visível. Sempre conferir os 3 containers Up.

---

## 6. SUBIR O APP + ACESSO

```bash
source .venv/bin/activate
# Rede interna (time acessa por http://IP-DA-VM:5000):
python -m waitress --listen=0.0.0.0:5000 app:app
# (se for só local + ngrok, usar 127.0.0.1:5000 e subir o ngrok: ngrok http 5000)
```
- **Acesso interno:** descobrir o IP da VM (`ip a`) e passar `http://IP-DA-VM:5000` pro time. Liberar a porta 5000 no firewall da VM pra rede interna (`sudo ufw allow from <rede> to any port 5000`).
- **Senha de acesso** = `APP_PASSWORD`.

---

## 7. SELENIUM HEADLESS — JÁ RESOLVIDO no código ✅

O `confio.py` já roda headless e com as flags certas de servidor Linux:
```python
# criar_driver() — navegador configuravel via SIGAVI_BROWSER (edge|chrome)
--headless=new --no-sandbox --disable-dev-shm-usage --disable-gpu
```
Só garantir no `.env`: `SIGAVI_BROWSER=chrome` e `SIGAVI_HEADLESS=true`. Nada mais a mexer.

---

## 8. RODAR 24/7 com systemd

Criar 2 units (app; e ngrok só se usar). Exemplo do app — `/etc/systemd/system/importador.service`:
```ini
[Unit]
Description=Importador Sigavi (Flask/waitress)
After=network.target docker.service

[Service]
User=<usuario>
WorkingDirectory=/home/<usuario>/importadorEmails_v2
ExecStart=/home/<usuario>/importadorEmails_v2/.venv/bin/python -m waitress --listen=0.0.0.0:5000 app:app
Restart=always
EnvironmentFile=/home/<usuario>/importadorEmails_v2/.env

[Install]
WantedBy=multi-user.target
```
```bash
sudo systemctl daemon-reload
sudo systemctl enable --now importador
sudo systemctl status importador
```
- **Evolution (Docker):** o compose já tem `restart: always` → containers voltam sozinhos no boot (Docker sobe no boot por padrão).
- VM fica **sempre ligada** (confirmado — tem nobreak).

---

## 9. CHECKLIST FINAL (validar antes de entregar pro time)

- [ ] Ubuntu Server 24.04 com as specs (8GB/4vCPU/40GB), Docker liberado
- [ ] venv criado e `pip install -r requirements.txt` OK
- [ ] Google Chrome instalado; `SIGAVI_BROWSER=chrome` e `SIGAVI_HEADLESS=true` no `.env`; Selenium roda headless
- [ ] 3 containers da Evolution "Up"; `curl localhost:8080` responde
- [ ] WhatsApp **pareado na VM** (QR com o cel do Erick); mensagem de teste chegou
- [ ] `.env` (raiz) e `evolution/.env` preenchidos; `APP_PASSWORD` forte
- [ ] App rodando via systemd (`Restart=always`); sobe no boot
- [ ] Firewall: porta 5000 liberada só pra rede interna; 8080 protegida
- [ ] Teste ponta-a-ponta: link → login → subir planilha → rodar → resultado + WhatsApp OK
- [ ] Reboot de teste: VM reinicia e tudo volta sozinho

---

## 10. TROUBLESHOOTING RÁPIDO

| Sintoma | Causa provável | Solução |
|---|---|---|
| WhatsApp não chega, job roda normal | Evolution/Docker fora do ar | `docker compose up -d` em `evolution/`; conferir 3 containers Up |
| Selenium falha ao abrir navegador | `SIGAVI_BROWSER` errado / Chrome não instalado / sem headless | conferir Chrome instalado, `SIGAVI_BROWSER=chrome`, `SIGAVI_HEADLESS=true` |
| `session not created` / driver | versão do driver | o Selenium Manager baixa sozinho; garantir internet na VM no 1º run |
| Flask não sobe | porta 5000 ocupada / erro no app | `sudo lsof -i:5000`; ver `journalctl -u importador -e` |
| Time não acessa o link interno | firewall / IP errado | `ip a` pega o IP; `ufw allow ... port 5000` |
| Login não fica logado | cookie secure sem HTTPS | `SESSION_COOKIE_SECURE=false` no acesso HTTP interno |

---

## 11. PORTAS E CAMINHOS DE REFERÊNCIA

- App Flask (waitress): **0.0.0.0:5000** (`/health` pra checar)
- Evolution API: **localhost:8080** (`/manager` = painel do WhatsApp)
- Pastas do projeto: `uploads/`, `backups/`, `resultados/`, `abertos/`, `evolution/`
- Logs do app (systemd): `journalctl -u importador -e`

> Dúvidas de regra de negócio / valores reais dos `.env` / branch certa: **perguntar ao Erick.**
