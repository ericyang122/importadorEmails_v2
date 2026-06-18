# Importador Sigavi — Automação de Leads (Abyara)

Ferramenta de **RPA (automação de processos)** para o CRM **Sigavi360 da Abyara**
(`abyara.sigavi360.com.br`). Ela lê uma planilha de leads e, de forma automática,
faz uma de duas coisas conforme o modo escolhido:

- **Consulta** → para cada e-mail da planilha, **busca o telefone** do lead dentro do Sigavi.
- **Cadastro** → **cadastra os leads** (que já têm telefone) no Sigavi, evitando duplicados.

Tudo isso roda por uma **interface web local** (feita pro time de Marketing usar sem mexer
em código) e, ao final, pode mandar um **resumo + as planilhas de resultado pelo WhatsApp**.

> 📄 Guia de operação e solução de problemas (passo a passo do dia a dia) está no Obsidian:
> `Eric/Projetos/Importador Sigavi - Operação e Troubleshooting`.

---

## 1. Como funciona (visão geral)

```
┌─────────────┐    sobe planilha .xlsx     ┌──────────────────┐
│  Navegador  │ ─────────────────────────► │   app.py (Flask) │  ← interface web (porta 5000)
│ (Marketing) │ ◄───── progresso/ETA ───── │  orquestra tudo  │
└─────────────┘                            └────────┬─────────┘
                                                    │ chama como subprocesso
                                                    ▼
                                          ┌────────────────────┐
                                          │     confio.py      │  ← o robô (Selenium + Edge)
                                          │ loga no Sigavi e    │
                                          │ consulta/cadastra   │
                                          └─────┬──────────┬────┘
                                                │          │ ao terminar
                          gera planilhas .xlsx  │          ▼
                                                │   ┌──────────────┐   ┌───────────────┐
                                                │   │ whatsapp.py  │──►│ Evolution API │ ← Docker (porta 8080)
                                                ▼   └──────────────┘   │  → WhatsApp   │
                                          resultados/ + backups/       └───────────────┘
```

**Fluxo resumido:** o usuário abre a tela → escolhe o modo → digita login/senha do Sigavi
→ sobe a planilha → o `app.py` dispara o `confio.py` → o robô loga no Sigavi e processa
linha a linha, reportando o progresso de volta pra tela → no fim, gera as planilhas de
resultado, salva um backup e (se configurado) manda tudo pelo WhatsApp.

---

## 2. Ferramentas usadas e POR QUÊ

| Ferramenta | Papel no projeto | Por que foi escolhida |
|-----------|------------------|------------------------|
| **Python 3.12** | Linguagem de tudo | Ecossistema forte pra automação e dados |
| **Flask** | Servidor web da interface (`app.py`) | Leve e simples pra subir uma tela local rápida |
| **Selenium 4** | Controla o navegador pra "robotizar" o Sigavi (`confio.py`) | O Sigavi não tem API pública; a única forma é simular um humano usando o site |
| **Microsoft Edge** | Navegador que o Selenium dirige | Já vem no Windows; o driver é baixado sozinho pelo Selenium Manager (sem instalar nada à mão) |
| **pandas + openpyxl** | Ler/escrever as planilhas `.xlsx` | Padrão de mercado pra Excel em Python |
| **requests** | Buscas rápidas por e-mail (consulta) e falar com a Evolution API | Mais rápido que abrir o navegador pra cada e-mail |
| **python-dotenv** | Carrega segredos do arquivo `.env` | Mantém senha/API key fora do código e fora do git |
| **Evolution API** (Docker) | Ponte com o WhatsApp | Permite mandar mensagem/arquivo pelo WhatsApp via API |
| **Docker** (Postgres + Redis + Evolution) | Roda a Evolution API e seus bancos | Sobe tudo isolado com um comando, sem instalar serviços no PC |

---

## 3. Estrutura de pastas

```
importadorEmails_v2/
├── app.py                ← servidor Flask (a interface web + orquestração dos jobs)
├── confio.py             ← o robô: Selenium/Edge que loga e consulta/cadastra no Sigavi
├── whatsapp.py           ← envio do resumo + planilhas via Evolution API (opcional)
├── corretores.json       ← mapeamento corretor → gerente (usado no cadastro)
├── requirements.txt      ← dependências Python
├── .env                  ← segredos (NÃO sobe pro git)
├── .env.example          ← modelo do .env pra copiar
├── templates/
│   └── index.html        ← a página da interface
├── static/
│   ├── app.js            ← lógica da tela (upload, prévia, progresso ao vivo)
│   └── app.css           ← estilo
├── evolution/
│   └── docker-compose.yml← sobe a Evolution API + Postgres + Redis (WhatsApp)
├── abertos/              ← (uso por linha de comando) coloque aqui os Excel a importar
├── uploads/              ← planilhas enviadas pela tela (temporário, por job)
├── resultados/           ← planilhas de resultado
└── backups/              ← 1 pasta por execução: entrada + resultados + progresso + log
```

---

## 4. Pré-requisitos e instalação

- **Windows** com **Python 3.12** e **Microsoft Edge** instalados.
- **Docker Desktop** (só se for usar a notificação por WhatsApp).

```bash
pip install -r requirements.txt
```

> O driver do Edge é gerenciado **automaticamente** pelo Selenium 4 (Selenium Manager).
> Não precisa baixar driver manualmente.

---

## 5. Configuração (`.env`)

Copie o `.env.example` para `.env` e preencha. **O `.env` nunca sobe pro git** (está no `.gitignore`).

### 5.1 Credenciais do Sigavi
```env
SIGAVI_LOGIN=seu.email@abyara.com.br
SIGAVI_SENHA=sua_senha
```
> Usadas **pela linha de comando**. Na **interface web**, o login/senha do Sigavi são
> digitados na própria tela e passados só para aquela execução (não ficam salvos no `.env`).

### 5.2 Tela web (`app.py`)
```env
APP_PASSWORD=uma_senha_forte          # senha que protege o acesso à tela (antes do form do Sigavi)
APP_SECRET_KEY=uma_chave_grande_aleatoria  # chave interna do Flask pra assinar a sessão do navegador
```
> ⚠️ Se `APP_PASSWORD` não for definida, a senha local padrão é `marketing123`.
> **Defina uma senha forte antes de expor a ferramenta pra outras pessoas.**

### 5.3 Desempenho (opcional)
```env
SIGAVI_CONSULTA_WORKERS=8       # buscas de e-mail simultâneas (baixe se o Sigavi limitar)
SIGAVI_CONSULTA_TENTATIVAS=3    # retentativas quando a resposta vem suspeita/vazia
SIGAVI_CONSULTA_BACKOFF=1.5     # pausa base (s) entre retentativas (cresce a cada tentativa)
SIGAVI_CADASTRO_DELAY=1.0       # fator de velocidade do cadastro (0.6 = mais rápido; suba se a tela travar)
```

### 5.4 WhatsApp (opcional — ver seção 9)
```env
EVOLUTION_API_URL=http://localhost:8080
EVOLUTION_API_KEY=sua_api_key
EVOLUTION_INSTANCE=nome_da_instancia
WHATSAPP_DESTINO=5511999998888   # número que recebe o resumo (com DDI, sem +)
```
> Deixe em branco para **desativar** o WhatsApp — o sistema roda normalmente sem ele.

---

## 6. Como rodar

### 6.1 Pela interface web (recomendado)
```bash
python app.py
```
Acesse **http://localhost:5000**, faça login com a `APP_PASSWORD`, e na tela:
1. Escolha **Somente consulta** ou **Somente cadastro**.
2. Digite **login e senha do Sigavi**.
3. Arraste a planilha `.xlsx` (drag & drop).
4. Confira a **prévia** (ver seção 7) e clique em iniciar.
5. Acompanhe o progresso ao vivo (barra com %/ETA, cards, "Últimas ações").
6. No fim: painel com resumo + botões de **download** das planilhas (e WhatsApp, se ativo).

> Botão **Parar e salvar**: pede parada segura e grava os resultados no próximo ponto de salvamento.

### 6.2 Pela linha de comando
```bash
# Consulta (busca telefone por e-mail)
python confio.py --mode consulta --excel ./abertos/arquivo.xlsx

# Cadastro (cadastra leads com telefone)
python confio.py --mode cadastro --excel ./abertos/arquivo.xlsx

# Navegador oculto (sem janela)
python confio.py --mode consulta --excel ./abertos/arquivo.xlsx --headless

# Escolher pasta de resultado / arquivos de progresso e parada
python confio.py --mode consulta --excel ./abertos/arquivo.xlsx \
  --result-dir ./resultados \
  --progress-file ./backups/teste/progresso.json \
  --stop-file ./backups/teste/parar.flag
```
**Pausar/retomar (CLI):** `Ctrl+C` salva o progresso; rodar de novo retoma de onde parou.

---

## 7. A prévia da planilha (validação por modo)

Ao subir a planilha, a tela mostra uma **prévia** com as primeiras linhas e valida as
colunas **de acordo com o modo** (cada modo precisa de coisas diferentes):

- **Consulta** → só precisa da coluna de **e-mail**. Com e-mail presente, mostra uma
  confirmação **verde** ("e-mails prontos pra buscar telefone"). Sem e-mail, avisa em amarelo.
- **Cadastro** → precisa de **Nome, Telefone, Corretor, Empreendimento**. Se faltar alguma,
  mostra aviso **amarelo** listando o que falta; se estiver tudo, confirmação **verde**.

O app reconhece variações de nome de cabeçalho (renomeadas em `COLUNAS_RENOMEAR` no `app.py`):

| Coluna | Nomes aceitos no cabeçalho |
|--------|----------------------------|
| Nome | `NOME`, `NOME COMPLETO`, `nome_cliente` |
| Telefone | `TELEFONE`, `celular` |
| Corretor | `CORRETOR DE ORIGEM`, `CORRETOR ORIGEM`, `corretor` |
| Empreendimento | qualquer coluna com `EMPREEND` no nome |

> Para reconhecer um cabeçalho novo, adicione o nome no dicionário `COLUNAS_RENOMEAR`.
> Mexeu no `app.js`/`app.css`? Dê **Ctrl+F5** no navegador (refresh forçado) pra não pegar o cache antigo.

---

## 8. Validação de login (proteção contra o "bug dos 51 erros")

Antes, se o login no Sigavi falhasse (senha errada, **sessão concorrente** aberta em outro
lugar, ou captcha), o robô **seguia mesmo assim** e gerava erro em **todas** as linhas.

Agora, logo após o login, o `confio.py` **abre a home e confere se a sessão está autenticada**:
- **Logou** → imprime `Login no Sigavi confirmado.` e segue.
- **Não logou** → **para na hora** com uma mensagem clara (credenciais / sessão concorrente /
  captcha) e encerra com erro, sem processar a planilha à toa.

---

## 9. Notificação por WhatsApp (Evolution API + Docker)

Ao final de cada execução, o sistema pode mandar um **resumo + as planilhas** pelo WhatsApp.
É **opcional**: sem as variáveis preenchidas (seção 5.4), nada é enviado e nada quebra.

### Como o WhatsApp é entregue
O `whatsapp.py` fala com a **Evolution API**, que roda em **Docker** na porta **8080** e é
quem está de fato conectada ao número de WhatsApp. Containers (em `evolution/docker-compose.yml`):

| Container | Imagem | Papel |
|-----------|--------|-------|
| `evolution_api` | `atendai/evolution-api:v2.1.1` | a API que envia as mensagens (porta 8080) |
| `evolution_postgres` | `postgres:15` | banco de dados da Evolution |
| `evolution_redis` | `redis:7` | cache/sessão da Evolution |

### Subir a Evolution
```bash
cd evolution
docker compose up -d
```
Os containers têm restart automático — uma vez no ar, voltam sozinhos quando o Docker abre.

> ⚠️ **Dependência importante:** se o **Docker estiver desligado**, a Evolution fica fora do
> ar e o envio **falha em silêncio** (o job roda e gera as planilhas, mas nada chega no
> WhatsApp e nenhum erro aparece na tela). Sintoma clássico: "job terminou ok mas o resumo
> não chegou no zap" → confira se o Docker está rodando (`docker ps` deve listar os 3
> containers; `curl http://localhost:8080` deve responder "Welcome to the Evolution API").

O resumo inclui: modo, duração, total processado e os contadores (encontrados/cadastrados,
não encontrados/duplicados, erros), seguido das planilhas `.xlsx` em anexo.
Compatível com Evolution API v2 (`/message/sendText` e `/message/sendMedia`).

---

## 10. Endpoints da interface (`app.py`)

| Método | Rota | O que faz |
|--------|------|-----------|
| GET | `/` | Página principal (tela de login ou o formulário) |
| POST | `/login` | Autentica com a `APP_PASSWORD` |
| POST | `/logout` | Encerra a sessão |
| POST | `/preview` | Lê a planilha e devolve a prévia + validação por modo (não executa nada) |
| POST | `/jobs` | Inicia uma execução (recebe modo, login Sigavi e planilha) |
| GET | `/jobs/<id>` | Status/progresso/log de um job (usado no polling ao vivo) |
| POST | `/jobs/<id>/stop` | Pede parada segura ("Parar e salvar") |
| GET | `/jobs/<id>/download` | Baixa todas as planilhas de resultado |
| GET | `/jobs/<id>/download/<file_id>` | Baixa uma planilha específica |
| GET | `/health` | Healthcheck simples (`{"ok": true}`) |

---

## 11. Resultados, backups e retomada

- **`resultados/`** e o painel final trazem as planilhas geradas:
  - **Consulta:** `encontrados`, `nao_encontrados`, `erros_consulta`
  - **Cadastro:** `cadastrados`, `duplicados`, `nao_cadastrados`, `erros_cadastro`
- **`uploads/`** é temporária (1 subpasta por job).
- **`backups/`** guarda 1 pasta por execução com: planilha de entrada, resultados,
  `progresso.json` e `log.txt`. **Preserve essa pasta** — é o que permite retomar e auditar.

---

## 12. Segurança

- O `.env` (com senhas e API key) **não vai pro git** (`.gitignore`).
- A pasta `evolution/` **não vai pro git** (o compose tem a API key embutida).
- Defina `APP_PASSWORD` e `APP_SECRET_KEY` fortes **antes de expor** a ferramenta.
- As credenciais do Sigavi digitadas na tela **não são gravadas** — valem só para aquela execução.

---

## 13. Roadmap

- [x] **nº 7** — botão "reprocessar erros" (rodar de novo só as linhas que falharam) ✅ feito (commit 8a96d3d: backend `/jobs/<id>/reprocess` + botão no front)
- [ ] **nº 3** — centralizar os XPaths do Sigavi num lugar só (robustez) — ⚠️ refatoração grande (62 XPaths), fazer COM teste contra o Sigavi
- [ ] **nº 4** — refatorar `confio.py` em funções
- [ ] **nº 8** — CSS responsivo: modelo pra celular (mobile-friendly)
- [ ] Subir a versão pra branch `main` (hoje o trabalho está em `teste-automacao`)
- [ ] Deploy: VM **Ubuntu Server** 24/7 (ver `DEPLOY_VM.md`) — antes era via ngrok
