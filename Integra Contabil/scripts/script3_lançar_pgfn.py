"""
script3_pgfn_sispar.py

Fluxo geral:

1) Lê LISTA PARCELAMENTOS.xlsx na PASTA MÃE e encontra linhas com "PGFN".
2) Para cada (CPF/CNPJ, número PGFN) único:
   - Acessa o site Sispar PGFN.
   - Preenche CPF/CNPJ e Número do Parcelamento.
   - Marca o "checkbox falso" via TAB/TAB/ENTER.
   - Clica em Consultar.
   - Na tabela de parcelas:
       * Encontra a linha do mês/ano atual (MM/AAAA), rolando a página
         para baixo se necessário.
       * Nessa linha:
           - Se existir ícone de imprimir com title "Não emitido":
                 clica nele, confirma emissão e depois clica no link do PDF.
           - Se NÃO existir impressora "Não emitido", mas existir ícone de
             alerta/"!": não clica em nada e registra esse CPF/CNPJ numa
             lista de conferência.
           - Se não tiver nem uma coisa nem outra: ignora.
   - Página de confirmação da emissão:
       * Clica no botão de confirmação ("Emitir"/"Confirmar"/"Gerar", etc.).
   - Página de EMISSÃO DE DOCUMENTO DE ARRECADAÇÃO:
       * Clica no link do PDF (id formResumoParcelamentoDarf:emitirDarf)
         para baixar o DARF.
       * O PDF é salvo na pasta PASTA_MAE/PDFs (mantendo o nome padrão do site).

Saídas extras:
- PDFs em: PASTA_MAE/PDFs
- CNPJs com "!" (sem impressora no mês atual), APÓS FILTRAR COM BASE NOS PDFs EXISTENTES:
    PASTA_MAE/pgfn_alertas_exclamacao.xlsx

Qualquer caso em que já exista algum PDF cujo conteúdo contenha o mesmo
CPF/CNPJ e o mesmo ano do MES_ANO do alerta é removido da planilha de alertas.
"""

import re
import time
import datetime as dt
from pathlib import Path
from typing import List, Dict, Tuple

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    WebDriverException,
)
from selenium.webdriver import ActionChains

# =========================
# CONFIG BÁSICA
# =========================

URL_SISPAR = (
    "https://sisparnet.pgfn.fazenda.gov.br/"
    "sisparInternet/internet/darf/consultaParcelamentoDarfInternet.xhtml"
)

NOME_ARQ_LISTA = "LISTA PARCELAMENTOS.xlsx"
NOME_AUTH = "auth_integra.py"  # só para localizar PASTA_MAE igual aos outros scripts


def encontrar_pasta_mae(
    nome_arquivo_lista: str = NOME_ARQ_LISTA,
    nome_auth: str = NOME_AUTH,
    max_niveis: int = 5,
) -> Path:
    pasta = Path(__file__).resolve().parent
    for _ in range(max_niveis):
        if (pasta / nome_arquivo_lista).exists() and (pasta / nome_auth).exists():
            return pasta
        pasta = pasta.parent
    raise FileNotFoundError(
        f'Não encontrei a pasta mãe contendo "{nome_arquivo_lista}" '
        f'e "{nome_auth}" nos {max_niveis} níveis acima.'
    )


PASTA_MAE = encontrar_pasta_mae()
ARQ_ENTRADA = PASTA_MAE / NOME_ARQ_LISTA

# pasta para salvar os DARFs baixados (dentro de Integra Contabil)
PASTA_DOWNLOAD = PASTA_MAE / "PDFs"
PASTA_DOWNLOAD.mkdir(parents=True, exist_ok=True)

ARQ_ALERTAS = PASTA_MAE / "pgfn_alertas_exclamacao.xlsx"

# =========================
# LEITURA DO EXCEL E EXTRAÇÃO PGFN
# =========================


def carregar_casos_pgfn() -> List[Dict[str, str]]:
    if not ARQ_ENTRADA.exists():
        raise FileNotFoundError(
            f'Arquivo "{ARQ_ENTRADA}" não encontrado. '
            f'Confirme a PASTA MÃE: {PASTA_MAE}'
        )

    df = pd.read_excel(ARQ_ENTRADA)

    col_doc = "CPF/CNPJ"
    col_parcelamento = "PARCELAMENTO"

    for col in (col_doc, col_parcelamento):
        if col not in df.columns:
            raise KeyError(
                f'A coluna "{col}" não foi encontrada em {ARQ_ENTRADA.name}. '
                f'Confirme que o layout não foi alterado.'
            )

    # normaliza CPF/CNPJ
    df[col_doc] = (
        df[col_doc].astype(str).str.replace(r"\D", "", regex=True).str.strip()
    )
    df[col_parcelamento] = df[col_parcelamento].astype(str)

    # só linhas que mencionam PGFN
    df_pgfn = df[df[col_parcelamento].str.contains("PGFN", case=False, na=False)].copy()

    casos_set: set[Tuple[str, str]] = set()
    casos: List[Dict[str, str]] = []

    for _, row in df_pgfn.iterrows():
        cpf_cnpj = str(row[col_doc]).strip()
        if not cpf_cnpj:
            continue

        texto_parc = str(row[col_parcelamento])

        # pega o número logo após "PGFN"
        m = re.search(r"PGFN[^0-9]*([0-9][0-9./-]*)", texto_parc, re.IGNORECASE)
        if not m:
            continue

        numero_pgfn = m.group(1).strip()
        numero_pgfn_limpo = numero_pgfn  # mantemos com / - se tiver

        chave = (cpf_cnpj, numero_pgfn_limpo)
        if chave in casos_set:
            continue

        casos_set.add(chave)
        casos.append(
            {
                "cpf_cnpj": cpf_cnpj,
                "parcelamento_pgfn": numero_pgfn_limpo,
                "descricao_parcelamento": texto_parc,
            }
        )

    return casos


# =========================
# AUTOMACAO SELENIUM
# =========================


def abrir_navegador() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    # configura diretório de download e PDF pra baixar direto
    prefs = {
        "download.default_directory": str(PASTA_DOWNLOAD),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        # força abrir PDF externamente (download em vez de visualizar no Chrome)
        "plugins.always_open_pdf_externally": True,
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)
    return driver


def clicar_botao_cookies_se_existir(driver: webdriver.Chrome) -> None:
    try:
        wait = WebDriverWait(driver, 5)
        botao = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//button[contains(., 'Permitir') or contains(., 'Rejeitar') "
                    "or contains(., 'Aceitar') or contains(., 'OK')]",
                )
            )
        )
        botao.click()
    except TimeoutException:
        pass
    except Exception:
        pass


def preencher_formulario_inicial(
    driver: webdriver.Chrome,
    cpf_cnpj: str,
    numero_parcelamento: str,
) -> None:
    """
    Preenche CPF/CNPJ, Número do Parcelamento e
    faz a sequência de teclas:
        TAB, TAB, ENTER
    com delay de 0,3s entre cada uma (via ActionChains),
    depois espera 3s e clica no botão 'Consultar'.
    """
    wait = WebDriverWait(driver, 20)

    # Campo CPF/CNPJ
    campo_cpf = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//label[contains(., 'CPF/CNPJ') or contains(., 'Número do CPF')]/following::input[1]",
            )
        )
    )
    campo_cpf.clear()
    campo_cpf.send_keys(cpf_cnpj)
    time.sleep(0.3)

    # Campo Número do Parcelamento
    campo_parcelamento = driver.find_element(
        By.XPATH,
        "//label[contains(., 'Número do Parcelamento')]/following::input[1]",
    )
    campo_parcelamento.clear()
    campo_parcelamento.click()  # garante o foco nesse campo
    campo_parcelamento.send_keys(numero_parcelamento)
    time.sleep(0.3)

    # Sequência via teclado (global) usando ActionChains: TAB, TAB, ENTER
    actions = ActionChains(driver)
    actions.send_keys(Keys.TAB).pause(0.3)
    actions.send_keys(Keys.TAB).pause(0.3)
    actions.send_keys(Keys.ENTER).pause(0.3)
    actions.perform()

    # Espera 3 segundos antes de clicar em Consultar
    time.sleep(3.0)

    # Agora clica no botão Consultar (id conhecido)
    try:
        botao_consultar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.ID, "formConsultaParcelamentoDarf:btnConsultar")
            )
        )
        botao_consultar.click()
    except TimeoutException:
        # fallback: tenta pelo texto, se der algum problema com o id
        try:
            botao_fallback = driver.find_element(
                By.XPATH,
                "//button[@id='formConsultaParcelamentoDarf:btnConsultar' or "
                "contains(normalize-space(.), 'Consultar')]",
            )
            botao_fallback.click()
        except Exception:
            pass


def clicar_botao_confirmar_emissao(driver: webdriver.Chrome) -> bool:
    """
    Página logo após clicar na impressora:
      - procura um botão/input/link que confirme a emissão
        (texto ou value contendo 'Emitir', 'Confirmar' ou 'Gerar').
      - clica UMA vez, se encontrar.

    Retorna True se conseguiu clicar, False caso contrário.
    """
    wait = WebDriverWait(driver, 20)

    xpaths_botoes = [
        # botões visíveis com esses textos
        "//button[contains(normalize-space(.),'Emitir') "
        "      or contains(normalize-space(.),'Confirmar') "
        "      or contains(normalize-space(.),'Gerar')]",
        # inputs submit/botão
        "//input[@type='submit' or @type='button'][ "
        " (contains(@value,'Emitir') or contains(@value,'Confirmar') or contains(@value,'Gerar'))]",
        # links
        "//a[contains(normalize-space(.),'Emitir') "
        "  or contains(normalize-space(.),'Confirmar') "
        "  or contains(normalize-space(.),'Gerar')]",
    ]

    for xp in xpaths_botoes:
        try:
            botao = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
            botao.click()
            time.sleep(2.5)  # tempo pra carregar a próxima tela
            return True
        except TimeoutException:
            continue
        except WebDriverException:
            continue

    # não encontrou nenhum botão claro de confirmação
    return False


def clicar_link_pdf_emissao(
    driver: webdriver.Chrome,
    cpf_cnpj: str,
    numero_parcelamento: str,
    mes_ano_label: str,
) -> None:
    """
    Na página "EMISSÃO DE DOCUMENTO DE ARRECADAÇÃO - INTERNET",
    tenta clicar no link do PDF (ícone + número) para baixar/abrir o DARF.

    Agora prioriza o link com id:
        formResumoParcelamentoDarf:emitirDarf

    Estratégia:
      1) Tenta pelo ID exato.
      2) Se não funcionar, tenta XPaths genéricos.
      3) Se ainda falhar, usa fallback com TAB/ENTER.
    """
    wait = WebDriverWait(driver, 30)
    clicou = False

    # 1) Tenta pelo ID específico
    try:
        link_pdf = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "formResumoParcelamentoDarf:emitirDarf")
            )
        )
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", link_pdf
            )
            time.sleep(0.5)
        except WebDriverException:
            pass
        link_pdf.click()
        time.sleep(3.0)
        clicou = True
    except TimeoutException:
        clicou = False
    except WebDriverException:
        clicou = False

    # 2) XPaths genéricos, se o ID não deu certo
    if not clicou:
        xpaths_links = [
            # link que contém img de pdf
            "//a[.//img[contains(translate(@src,'PDF','pdf'),'pdf')]]",
            # link cujo href parece pdf/darf
            "//a[contains(translate(@href,'PDF','pdf'),'.pdf') "
            " or contains(translate(@href,'DARF','darf'),'darf')]",
            # link cujo texto é praticamente só números (ex.: 7172534290179884)
            "//a[translate(normalize-space(string(.)),'0123456789','')='' "
            " and string-length(normalize-space(string(.)))>=5]",
            # fallback: primeiro link que não seja 'Voltar'
            "//a[not(contains(normalize-space(.),'Voltar'))]",
        ]

        for xp in xpaths_links:
            try:
                link_pdf = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
                try:
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block: 'center'});",
                        link_pdf,
                    )
                    time.sleep(0.5)
                except WebDriverException:
                    pass
                link_pdf.click()
                time.sleep(3.0)
                clicou = True
                break
            except TimeoutException:
                continue
            except WebDriverException:
                continue

    # 3) Fallback com TAB/ENTER, se nada funcionou
    if not clicou:
        try:
            actions = ActionChains(driver)
            actions.send_keys(Keys.TAB).pause(0.3)
            actions.send_keys(Keys.TAB).pause(0.3)
            actions.send_keys(Keys.ENTER).pause(0.3)
            actions.perform()
            time.sleep(3.0)
            clicou = True
        except WebDriverException:
            pass

    # Não renomeamos nenhum PDF aqui — apenas deixamos o download acontecer.


def selecionar_parcela_mes_atual(
    driver: webdriver.Chrome,
    cpf_cnpj: str,
    numero_parcelamento: str,
    lista_alertas: List[Dict[str, str]],
) -> None:
    """
    Na tela de resumo, acha a linha que contém o mês/ano atual (MM/AAAA),
    rolando a página para baixo se necessário, e tenta achar dentro dessa linha
    o ícone de impressora:

      <img id="formResumoParcelamentoDarf:tabaplicacaoparcelamentoparcela:XX:NNimprimirDarf"
           title="Não emitido" ...>

    Comportamento:

      - Se encontrar impressora "Não emitido": clica, confirma emissão
        e tenta baixar o PDF.

      - Se NÃO encontrar impressora "Não emitido", mas encontrar ícone de
        alerta / '!' nessa mesma linha: NÃO clica em nada e registra
        esse CPF/CNPJ em lista_alertas, junto com o mês/ano.

      - Se não achar nenhuma dessas coisas: ignora silenciosamente.
    """
    hoje = dt.date.today()
    mes_ano_label = hoje.strftime("%m/%Y")  # ex.: "12/2025"

    wait = WebDriverWait(driver, 20)

    # garante que a tabela de parcelas apareceu
    tabela_xpath = (
        "//tbody[contains(@id,'tabaplicacaoparcelamentoparcela') or "
        "       contains(@id,'tabaplicacaoparcelamentoparcela_data')]"
    )
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, tabela_xpath)))
    except TimeoutException:
        return

    # tenta achar a linha rolando a tela para baixo, aos poucos
    linha = None
    row_xpath = (
        tabela_xpath
        + f"//tr[.//td[contains(normalize-space(.), '{mes_ano_label}')]]"
    )

    # começa do topo da página
    try:
        driver.execute_script("window.scrollTo(0, 0);")
    except WebDriverException:
        pass

    for _ in range(12):  # até ~12 passos de scroll
        try:
            linha = driver.find_element(By.XPATH, row_xpath)
            break
        except NoSuchElementException:
            try:
                driver.execute_script("window.scrollBy(0, 500);")
            except WebDriverException:
                pass
            time.sleep(0.4)

    if linha is None:
        # não achou linha do mês atual, não faz nada
        return

    # traz a linha pro centro da tela (para garantir clicabilidade)
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", linha
        )
        time.sleep(0.5)
    except WebDriverException:
        pass

    # primeiro tenta achar impressora "Não emitido"
    img_imprimir = None
    try:
        img_imprimir = linha.find_element(
            By.XPATH,
            ".//img[contains(@id,'imprimirDarf') and "
            "(contains(@title,'Não emitido') or contains(@title,'Nao emitido'))]",
        )
    except NoSuchElementException:
        img_imprimir = None

    if img_imprimir is None:
        # não há impressora - procura ícone de alerta / "!"
        try:
            linha.find_element(
                By.XPATH,
                ".//img[contains(@src,'alert') or contains(@src,'excl') or "
                "      contains(@alt,'!')   or contains(@title,'!')]",
            )
            # se chegou aqui, achou ícone de alerta
            lista_alertas.append(
                {
                    "CPF_CNPJ": cpf_cnpj,
                    "PARCELAMENTO_PGFN": numero_parcelamento,
                    "MES_ANO": mes_ano_label,
                    "MOTIVO": "Ícone de alerta/! em vez de impressora para o mês atual",
                }
            )
        except NoSuchElementException:
            # nem impressora, nem alerta -> ignora
            pass
        return

    # se temos impressora "Não emitido", tenta clicar
    try:
        img_imprimir.click()
    except WebDriverException:
        return

    # espera a página de confirmação carregar e tenta clicar no botão de emissão
    time.sleep(2.0)
    confirmou = clicar_botao_confirmar_emissao(driver)
    if not confirmou:
        return

    # depois da confirmação, deve aparecer a página com link do PDF
    clicar_link_pdf_emissao(
        driver, cpf_cnpj, numero_parcelamento, mes_ano_label
    )


def processar_parcelamento_pgfn(
    driver: webdriver.Chrome,
    cpf_cnpj: str,
    numero_parcelamento: str,
    lista_alertas: List[Dict[str, str]],
) -> None:
    driver.get(URL_SISPAR)
    clicar_botao_cookies_se_existir(driver)

    preencher_formulario_inicial(driver, cpf_cnpj, numero_parcelamento)

    # espera o resultado carregar
    time.sleep(5)

    selecionar_parcela_mes_atual(
        driver, cpf_cnpj, numero_parcelamento, lista_alertas
    )
    time.sleep(2)


def filtrar_alertas_pelos_pdfs(alertas: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """
    Recebe a lista de alertas (cada um com CPF_CNPJ, PARCELAMENTO_PGFN, MES_ANO)
    e remove qualquer alerta para o qual já exista algum PDF em PASTA_DOWNLOAD
    cujo conteúdo contenha:

        - o mesmo CPF/CNPJ (apenas dígitos) E
        - o mesmo ano de MES_ANO.

    Ou seja: se já existe DARF para aquele CNPJ naquele ano, supomos que
    a pendência já foi tratada em outro momento.
    """

    if not alertas:
        return []

    pdf_paths = list(PASTA_DOWNLOAD.glob("*.pdf"))
    if not pdf_paths:
        return alertas  # não há PDFs pra comparar

    # tenta importar PyPDF2; se não tiver, devolve lista original
    try:
        import PyPDF2  # type: ignore
    except ImportError:
        print(
            "\n[AVISO] PyPDF2 não está instalado. "
            "Não foi possível filtrar os alertas com base nos PDFs."
        )
        return alertas

    # monta índice: para cada PDF, string com apenas dígitos presente no texto
    pdf_digits_list: List[str] = []
    for pdf_path in pdf_paths:
        try:
            reader = PyPDF2.PdfReader(str(pdf_path))
            texto = ""
            # lê só as primeiras páginas (normalmente 1) por segurança
            for i, page in enumerate(reader.pages):
                if i >= 3:
                    break
                try:
                    texto += page.extract_text() or ""
                except Exception:
                    continue
            digitos = re.sub(r"\D", "", texto)
            pdf_digits_list.append(digitos)
        except Exception:
            pdf_digits_list.append("")

    alertas_filtrados: List[Dict[str, str]] = []

    for a in alertas:
        cpf = re.sub(r"\D", "", str(a.get("CPF_CNPJ", "")).strip())
        mes_ano = str(a.get("MES_ANO", "")).strip()
        # usamos só o ano como filtro adicional
        ano = mes_ano[-4:] if len(mes_ano) >= 4 else ""

        if not cpf:
            # se não tiver CPF/CNPJ, mantém alerta
            alertas_filtrados.append(a)
            continue

        encontrado = False
        for digitos_pdf in pdf_digits_list:
            if not digitos_pdf:
                continue
            if cpf in digitos_pdf and (not ano or ano in digitos_pdf):
                encontrado = True
                break

        if not encontrado:
            alertas_filtrados.append(a)
        # se encontrado == True, removemos esse alerta (já existe DARF em PDF)

    return alertas_filtrados


def main():
    casos = carregar_casos_pgfn()

    if not casos:
        print(
            f"Nenhuma linha com 'PGFN' encontrada na coluna PARCELAMENTO de {ARQ_ENTRADA}."
        )
        return

    print("Casos PGFN encontrados (CPF_CNPJ, número do parcelamento):")
    for c in casos:
        print(f"  - {c['cpf_cnpj']} | {c['parcelamento_pgfn']}")

    driver = abrir_navegador()
    alertas: List[Dict[str, str]] = []

    try:
        for caso in casos:
            cpf_cnpj = caso["cpf_cnpj"]
            numero_parcelamento = caso["parcelamento_pgfn"]

            print(
                f"\n=== CPF/CNPJ {cpf_cnpj} – parcelamento PGFN {numero_parcelamento} ==="
            )

            processar_parcelamento_pgfn(
                driver, cpf_cnpj, numero_parcelamento, alertas
            )
            time.sleep(1)

    finally:
        print("\nFechando navegador...")
        driver.quit()

    # Filtra alertas usando os PDFs já existentes
    alertas_filtrados = filtrar_alertas_pelos_pdfs(alertas)

    # Salva lista final de CNPJs com ícone de alerta/! e sem PDF correspondente
    if alertas_filtrados:
        df_alertas = pd.DataFrame(alertas_filtrados).drop_duplicates()
        df_alertas.to_excel(ARQ_ALERTAS, index=False)
        print(f"\nLista de CNPJs com '!' (sem PDF correspondente) salva em: {ARQ_ALERTAS}")
    else:
        print("\nNenhum caso pendente com ícone '!' após conferir PDFs existentes.")


if __name__ == "__main__":
    main()
