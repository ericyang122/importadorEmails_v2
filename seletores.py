"""XPaths e seletores do Sigavi, centralizados num lugar so.

Quando o Sigavi muda a tela e a automacao quebra, o conserto e AQUI — nao
espalhado pelo confio.py. Cada entrada e uma tupla `(By, valor)` pronta pros
helpers (wait_visible / safe_click / ...). Listas = candidatos em ordem de
prioridade (fallback): o helper tenta do primeiro ao ultimo, entao o XPath
absoluto (fragil) fica por ultimo, depois dos seletores por texto/classe.
"""
from selenium.webdriver.common.by import By


class Login:
    """Tela de login (/Acesso/Login)."""
    USUARIO = (By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[1]/div/input")
    SENHA   = (By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[2]/div/input")
    ENTRAR  = (By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[3]/div/button")


class Combo:
    """Combos genericos (select2 / Kendo) — por classe, mais robustos a layout."""
    INPUT_ABERTO = [
        (By.XPATH, "//div[contains(@class, 'select2-drop') and not(contains(@style, 'display: none'))]//input"),
        (By.XPATH, "//span[contains(@class, 'select2-container--open')]//input"),
        (By.XPATH, "//input[contains(@class, 'select2-input') or contains(@class, 'select2-search__field')]"),
    ]
    OPCOES = [
        (By.XPATH, "//*[contains(@class, 'select2-result-label')]"),
        (By.XPATH, "//*[contains(@class, 'select2-results__option')]"),
        (By.XPATH, "//*[contains(@class, 'select2-result') and not(contains(@class, 'select2-searching'))]"),
        (By.XPATH, "//*[@role='option']"),
    ]
    # XPath RELATIVO (usado em element.find_elements dentro do dropdown aberto)
    DROPDOWN_CANDIDATOS = ".//*[self::li or self::div or self::span or self::a][normalize-space()]"
    # CSS dos <li> de um combo Kendo aberto
    KENDO_ITENS = 'ul.k-list li, li.k-list-item, li.k-item, ul[role="listbox"] li'


class BuscaFac:
    """Tela de Pesquisa de FAC (/CRM/Fac) — busca e verificacao de telefone."""
    BUSCAR = [
        (By.XPATH, "//button[contains(normalize-space(.), 'Buscar')]"),
        (By.XPATH, "//a[contains(normalize-space(.), 'Buscar')]"),
        (By.XPATH, "//*[self::button or self::a][contains(@title, 'Buscar')]"),
        (By.XPATH, "//*[self::button or self::a][contains(@class, 'btn') and .//*[contains(@class, 'search')]]"),
        (By.XPATH, "/html/body/section/section/div/div/div[2]/div/div[1]/form/div[4]/div/button[2]"),
    ]
    RESULTADO_LINHAS = (By.XPATH, "//div[contains(., 'RESULTADO DA BUSCA')]/following::table[1]//tbody/tr")
    # campo de telefone da busca via XPath absoluto (o cadastro usa este direto)
    TELEFONE_INPUT_ABS = (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[1]/form/div[2]/div/div/div/div[10]/div[4]/input')
    TELEFONE = [
        (By.XPATH, "//input[contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'telefone')]"),
        (By.XPATH, "//input[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'telefone')]"),
        (By.XPATH, "//input[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'telefone')]"),
        (By.XPATH, "//label[contains(normalize-space(.), 'Telefone')]/following::input[1]"),
        (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[1]/form/div[3]/div/div[3]/input'),
        TELEFONE_INPUT_ABS,
    ]


class Cadastro:
    """Tela de cadastro de FAC (/CRM/Fac/Cadastro) e seu modal de Imovel/Origem."""
    NOME = (By.ID, 'Nome')
    # grade de resultado da busca de duplicidade (antes de cadastrar)
    RESULTADO_BUSCA_DUP = (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[2]/div[2]/div[2]/div/div[4]/table/tbody/tr')
    PRIMEIRA_TD = './td[1]'  # RELATIVO a uma linha <tr> do resultado

    ABRIR_TELEFONES     = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/div/a')
    CELULAR_COMBO       = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[1]/span[1]/span/span[1]')
    TELEFONE_GRID_INPUT = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[3]/input')
    ADD_TELEFONE        = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[4]/a[1]/span')

    CANAL_COMBO    = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[1]/span[2]/span/span[1]')
    MIDIA_COMBO    = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[2]/span[2]/span/span[1]')
    EQUIPE_COMBO   = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[1]/span[1]/span/span[1]')
    CORRETOR_COMBO = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/span[1]/span/span[1]')

    ABRIR_MODAL_IMOVEL = (By.XPATH, "/html/body/div[2]/form/div[3]/div/div/div[1]/div[2]/div[2]/a/span")
    MODAL_CONTAINER    = (By.XPATH, "/html/body/div[2]/div[2]/div/div")
    MODAL_OPCAO_TIPO   = (By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[1]/div/div/label[2]")
    EMPREENDIMENTO_INPUT = [
        (By.XPATH, "//input[contains(translate(@placeholder, 'EMPRENDIMTO', 'emprendimto'), 'empreendimento')]"),
        (By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/input"),
        (By.XPATH, "/html/body/div[2]/div[2]/div/div//input[@type='text']"),
    ]
    MODAL_BUSCAR = [
        (By.XPATH, "/html/body/div[2]/div[2]/div/div//button[contains(normalize-space(.), 'Buscar')]"),
        (By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/button"),
    ]
    IMOVEL_ORIGEM_COMANDO = (By.CSS_SELECTOR, "#dvImovelOrigemComando a")

    CONFIRMAR_FORM    = (By.XPATH, "/html/body/div[2]/form/div[1]/div/div[1]/button[2]")
    POPUP_DUPLICIDADE = (By.XPATH, '//*[@id="popVerificaDuplicidade"]/div/div/div[3]/button')
    SALVAR            = (By.XPATH, '//*[@id="cmdSalva"]')
