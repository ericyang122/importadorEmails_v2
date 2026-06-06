# Automação de Importação — Sigavi Abyara

Script Python que lê leads de uma planilha Excel e os cadastra automaticamente no CRM Sigavi (abyara.sigavi360.com.br) via Selenium. Para leads sem telefone, busca o número diretamente no Sigavi pelo email do lead.

---

## Estrutura de pastas

```
automa-o_abyara/
├── abertos/          ← coloque aqui os Excel a importar
├── resultados/       ← resultados das buscas de email são salvos aqui
├── confio.py         ← script principal
├── corretores.json   ← mapeamento corretor → gerente
├── .env              ← login e senha (não subir pro git)
└── README.md
```

---

## Instalação

Requer Python 3.9+ e Google Chrome instalado.

```bash
pip install -r requirements.txt
```

O ChromeDriver é gerenciado automaticamente pelo Selenium 4+.

---

## Configuração

### 1. Credenciais (`.env`) — OBRIGATÓRIO
O arquivo `.env` **não está incluso no repositório** por segurança. Você precisa criá-lo manualmente na raiz do projeto antes de rodar o script.

Crie um arquivo chamado `.env` com o conteúdo:
```
SIGAVI_LOGIN=seu.email@abyara.com.br
SIGAVI_SENHA=sua_senha
```

> Sem esse arquivo o script não consegue fazer login e vai travar na inicialização.

### 2. Planilha
Coloque o arquivo Excel na pasta `abertos/` e aponte para ele no início do `confio.py`:
```python
arquivo_excel = './abertos/nome_do_arquivo.xlsx'
```

A planilha deve ter (ou ser renomeável para) as colunas:
| Coluna original | Usado como |
|---|---|
| `NOME COMPLETO` | Nome do lead |
| `TELEFONE` | Telefone (se vazio, busca pelo email) |
| `E-MAIL` ou `Email` | Email (usado para buscar telefone) |
| `CORRETOR DE ORIGEM` | Corretor responsável |
| `TIPO PLANTAO` | Mídia/canal de origem |

---

## Uso

### Somente consulta

Busca telefone por email e nao cadastra nada no Sigavi:

```bash
python confio.py --mode consulta --excel ./abertos/abertos_cora.xlsx
```

Relatorios gerados:
- `resultado_<arquivo>_encontrados.xlsx`
- `resultado_<arquivo>_nao_encontrados.xlsx`
- `resultado_<arquivo>_erros_consulta.xlsx`

### Somente cadastro

Usa telefones ja existentes na planilha para verificar duplicidade e cadastrar leads no Sigavi:

```bash
python confio.py --mode cadastro --excel ./abertos/abertos_cora.xlsx
```

Relatorios gerados:
- `resultado_<arquivo>_cadastrados.xlsx`
- `resultado_<arquivo>_duplicados.xlsx`
- `resultado_<arquivo>_nao_cadastrados.xlsx`
- `resultado_<arquivo>_erros_cadastro.xlsx`

### Pausar e retomar
- Pressione `Ctrl+C` para pausar — o progresso é salvo automaticamente
- Rode o script novamente para retomar de onde parou

---

## Arquivos que precisam ir junto (outro computador)

| Arquivo | Obrigatório |
|---|---|
| `confio.py` | Sim |
| `corretores.json` | Sim (para cadastro) |
| `.env` | **Criar manualmente** — não está no repositório |
| `abertos/` | Sim (com o Excel) |
| `resultados/` | Criar vazia no destino |

---

## Interface local Flask

Para rodar a tela local:

```bash
python app.py
```

Acesse:

```text
http://127.0.0.1:5000
```

Senha local padrao da ferramenta:

```text
marketing123
```

Antes de publicar para outras pessoas, defina uma senha propria no `.env`:

```env
APP_PASSWORD=uma_senha_forte
APP_SECRET_KEY=uma_chave_grande_aleatoria
```

`APP_PASSWORD` e a senha da tela do importador. Ela protege o acesso antes de aparecer o formulario do Sigavi.
`APP_SECRET_KEY` e uma chave interna do Flask para assinar a sessao do navegador; use um texto grande e aleatorio e nao compartilhe.

Na tela, o usuario escolhe `Somente consulta` ou `Somente cadastro`, informa login/senha do Sigavi e envia a planilha `.xlsx`.
Essas credenciais nao sao gravadas em `.env`; elas sao passadas apenas para a execucao do `confio.py`.
A interface tambem tem o botao `Parar e salvar`, que pede parada segura e grava os resultados no proximo ponto de salvamento.

Cada execucao cria um backup em `backups/` com:
- planilha de entrada;
- relatorios gerados;
- `progresso.json`;
- `log.txt`.

A pasta `uploads/` e temporaria. A pasta `backups/` deve ser preservada para retomada e auditoria.

O script tambem aceita planilha por parametro:

```bash
python confio.py --excel ./abertos/abertos_cora.xlsx
```

Para rodar o navegador em modo oculto:

```bash
python confio.py --excel ./abertos/abertos_cora.xlsx --headless
```

Para escolher a pasta de resultado:

```bash
python confio.py --excel ./abertos/abertos_cora.xlsx --result-dir ./resultados
```

Para escolher arquivo de progresso e parada segura:

```bash
python confio.py --mode consulta --excel ./abertos/abertos_cora.xlsx --progress-file ./backups/teste/progresso.json --stop-file ./backups/teste/parar.flag
```

---

## Ajustes de desempenho (`.env`)

A consulta por email roda em paralelo e o cadastro tem velocidade calibravel:

```env
SIGAVI_CONSULTA_WORKERS=8       # buscas de email simultaneas (baixe se houver throttle)
SIGAVI_CONSULTA_TENTATIVAS=3    # retentativas em resposta suspeita do Sigavi
SIGAVI_CONSULTA_BACKOFF=1.5     # pausa base (s) entre retentativas
SIGAVI_CADASTRO_DELAY=1.0       # fator de velocidade do cadastro (0.6 = mais rapido)
```

`SIGAVI_CADASTRO_DELAY` multiplica todas as pausas do cadastro. Calibre com a tela
visivel e uma planilha pequena: 1.0 e o padrao seguro; abaixe ate o ponto que ainda
cadastra corretamente.

---

## Notificacao por WhatsApp (Evolution API)

Ao terminar uma execucao, o sistema pode enviar um resumo + as planilhas de
resultado por WhatsApp. E **opcional**: sem configuracao, nada e enviado.

Para ativar, preencha no `.env`:

```env
EVOLUTION_API_URL=https://sua-evolution.com   # URL base da Evolution API
EVOLUTION_API_KEY=sua_api_key                 # apikey da instancia
EVOLUTION_INSTANCE=nome_da_instancia          # instancia conectada ao WhatsApp
WHATSAPP_DESTINO=5511999998888                # numero que recebe (com DDI, sem +)
```

O resumo inclui modo, duracao, processados e os contadores (encontrados/cadastrados,
nao encontrados/duplicados e erros), seguido das planilhas `.xlsx` em anexo.
Compativel com Evolution API v2 (`/message/sendText` e `/message/sendMedia`).
