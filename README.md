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
pip install selenium pandas requests openpyxl python-dotenv
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

```bash
python confio.py
```

O script vai:
1. Abrir o Chrome e fazer login no Sigavi
2. Para cada lead da planilha:
   - Se tem telefone → verifica duplicidade e cadastra
   - Se não tem telefone → busca no Sigavi pelo email
3. Ao final, salva `resultados/resultado_<arquivo>_1.xlsx` com duas abas:
   - **Com Telefone** — emails que tiveram telefone encontrado
   - **Todos** — todos os emails buscados

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
