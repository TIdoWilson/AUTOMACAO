# Baixador de Cota Unica - Francisco Beltrao

# =========================
# Configuracoes
# =========================

import csv
import os
import re
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime

try:
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
    from playwright.sync_api import sync_playwright
except Exception as exc:
    print("ERRO: dependencias nao encontradas.")
    print("Instale: pip install playwright")
    print("Depois rode: playwright install")
    print(f"Detalhe: {exc}")
    raise


URL_REIMPRESSAO = (
    "https://franciscobeltraopr.equiplano.com.br:7035/contribuinte/#/stmCarneAE/reimpressao"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CAMINHO_LISTA_EMPRESAS = os.path.join(BASE_DIR, "Lista_Empresas_Francisco_Beltrao.csv")
PASTA_SAIDA = os.path.join(BASE_DIR, "Cotas")
CAMINHO_LOG_ERROS = os.path.join(BASE_DIR, "Lista_Erros_Cotas.csv")
HEADLESS = False
TIMEOUT_PADRAO_MS = 30000
TEMPO_ENTRE_EMPRESAS_S = 1
TEMPO_AGUARDAR_APOS_CNPJ_MS = 2500
TIMEOUT_CARREGAMENTO_COTAS_MS = 60000
VALOR_MARCADOR_MES = "OK"

# Preencher com os seletores reais quando voce mapear o HTML do site.
SELETORES = {
    "campo_cnpj": "#cnpj",
    "botao_pesquisar": "TODO_BOTAO_PESQUISAR",
    "linhas_cota_unica": "tr:has-text('1 Parcela')",
    "botao_imprimir_linha": "button#imprimir-carne",
    "alerta_erro": "div.alert, div[role='alert'], .toast, .swal2-popup",
}


# =========================
# Modelos e utilitarios
# =========================


@dataclass
class Empresa:
    nome: str
    cnpj: str
    idx_linha_csv: int


def normalizar_cnpj(cnpj_bruto: str) -> str:
    return re.sub(r"\D", "", cnpj_bruto or "")


def normalizar_texto(texto: str) -> str:
    texto = texto or ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().lower()


def cnpj_valido(cnpj: str) -> bool:
    return len(cnpj) == 14 and cnpj.isdigit()


def limpar_nome_pasta(nome: str) -> str:
    nome = (nome or "").strip()
    nome = re.sub(r'[<>:"/\\|?*]', "_", nome)
    nome = re.sub(r"\s+", " ", nome)
    return nome or "SEM_NOME"


def nome_coluna_mes(data_ref: datetime) -> str:
    meses = [
        "jan",
        "fev",
        "mar",
        "abr",
        "mai",
        "jun",
        "jul",
        "ago",
        "set",
        "out",
        "nov",
        "dez",
    ]
    return f"{meses[data_ref.month - 1]} {data_ref.year}"


def _detectar_coluna_cnpj(cabecalho: list[str]) -> int:
    for i, col in enumerate(cabecalho):
        if "cnpj" in (col or "").strip().lower():
            return i
    return 0


def _detectar_coluna_nome(cabecalho: list[str]) -> int:
    # Regra solicitada: nome na coluna B.
    if len(cabecalho) > 1:
        return 1
    for i, col in enumerate(cabecalho):
        valor = (col or "").strip().lower()
        if "razao" in valor or "social" in valor or "nome" in valor:
            return i
    return 0


def _ler_csv_com_fallback(caminho: str) -> tuple[list[list[str]], str]:
    encodings_teste = ["utf-8-sig", "cp1252", "latin-1"]
    ultimo_erro: Exception | None = None

    for enc in encodings_teste:
        try:
            with open(caminho, "r", encoding=enc, newline="") as arquivo:
                leitor = csv.reader(arquivo, delimiter=";")
                linhas = [linha for linha in leitor if linha]
                return linhas, enc
        except UnicodeDecodeError as exc:
            ultimo_erro = exc
            continue

    print(f"ERRO: nao foi possivel ler o CSV por encoding: {caminho}")
    if ultimo_erro:
        print(f"Detalhe: {ultimo_erro}")
    sys.exit(1)


def carregar_empresas(caminho: str) -> tuple[list[Empresa], list[list[str]], int, str]:
    if not os.path.isfile(caminho):
        print(f"ERRO: lista de empresas nao encontrada: {caminho}")
        sys.exit(1)

    linhas, encoding_csv = _ler_csv_com_fallback(caminho)

    if not linhas:
        print("ERRO: CSV vazio.")
        sys.exit(1)

    cabecalho = linhas[0]
    idx_cnpj = _detectar_coluna_cnpj(cabecalho)
    idx_nome = _detectar_coluna_nome(cabecalho)

    empresas: list[Empresa] = []
    for idx_linha in range(1, len(linhas)):
        linha = linhas[idx_linha]
        cnpj_bruto = linha[idx_cnpj].strip() if idx_cnpj < len(linha) else ""
        nome = linha[idx_nome].strip() if idx_nome < len(linha) else ""

        cnpj = normalizar_cnpj(cnpj_bruto)
        if not cnpj_valido(cnpj):
            print(f"AVISO: linha {idx_linha + 1} ignorada (CNPJ invalido).")
            continue
        if not nome:
            nome = cnpj

        empresas.append(Empresa(nome=nome, cnpj=cnpj, idx_linha_csv=idx_linha))

    if not empresas:
        print("ERRO: nenhuma empresa valida encontrada na lista.")
        sys.exit(1)

    return empresas, linhas, idx_cnpj, encoding_csv


def garantir_coluna_mes(linhas_csv: list[list[str]], coluna_mes: str) -> int:
    cabecalho = linhas_csv[0]
    for i, col in enumerate(cabecalho):
        if (col or "").strip().lower() == coluna_mes.lower():
            return i

    cabecalho.append(coluna_mes)
    for idx in range(1, len(linhas_csv)):
        linhas_csv[idx].append("")
    return len(cabecalho) - 1


def salvar_csv(caminho: str, linhas_csv: list[list[str]], encoding_csv: str) -> None:
    with open(caminho, "w", encoding=encoding_csv, newline="") as arquivo:
        escritor = csv.writer(arquivo, delimiter=";")
        escritor.writerows(linhas_csv)


def registrar_mes_processado(
    linhas_csv: list[list[str]],
    idx_coluna_mes: int,
    empresa: Empresa,
    valor: str,
) -> None:
    linha = linhas_csv[empresa.idx_linha_csv]
    while len(linha) <= idx_coluna_mes:
        linha.append("")
    linha[idx_coluna_mes] = valor


def preparar_pasta_empresa(nome_empresa: str) -> str:
    pasta_empresa = os.path.join(PASTA_SAIDA, limpar_nome_pasta(nome_empresa))
    os.makedirs(pasta_empresa, exist_ok=True)
    return pasta_empresa


def registrar_erro_empresa(empresa: Empresa, motivo: str) -> None:
    arquivo_existe = os.path.isfile(CAMINHO_LOG_ERROS)
    with open(CAMINHO_LOG_ERROS, "a", encoding="utf-8-sig", newline="") as arquivo:
        writer = csv.writer(arquivo, delimiter=";")
        if not arquivo_existe:
            writer.writerow(["data_hora", "cnpj", "empresa", "motivo"])
        writer.writerow(
            [
                datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                empresa.cnpj,
                empresa.nome,
                motivo,
            ]
        )


# =========================
# Automacao web (esqueleto)
# =========================


def salvar_print_erro(page, pasta_empresa: str) -> None:
    caminho_print = os.path.join(pasta_empresa, "ERRO.png")
    try:
        page.screenshot(path=caminho_print, full_page=True)
        print(f"Print de erro salvo: {caminho_print}")
    except Exception as exc:
        print(f"AVISO: nao foi possivel salvar print de erro. Detalhe: {exc}")


def abrir_site(page) -> None:
    page.goto(URL_REIMPRESSAO, wait_until="domcontentloaded")
    page.set_default_timeout(TIMEOUT_PADRAO_MS)


def detectar_erro_debito(page) -> str | None:
    try:
        alertas = page.locator(SELETORES["alerta_erro"])
        total_alertas = alertas.count()
        for i in range(total_alertas):
            texto = alertas.nth(i).inner_text(timeout=500) or ""
            texto_norm = normalizar_texto(texto)
            if "debito" in texto_norm and "cnpj" in texto_norm:
                return texto.strip()
    except Exception:
        pass
    return None


def preencher_cnpj_e_pesquisar(page, cnpj: str) -> tuple[bool, str]:
    # Fluxo solicitado: digita CNPJ e envia Enter no proprio campo.
    page.click(SELETORES["campo_cnpj"])
    page.fill(SELETORES["campo_cnpj"], cnpj)
    page.locator(SELETORES["campo_cnpj"]).press("Enter")

    botao_pesquisar = (SELETORES.get("botao_pesquisar") or "").strip()
    if botao_pesquisar and not botao_pesquisar.startswith("TODO_"):
        page.click(botao_pesquisar)

    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightTimeoutError:
        pass
    page.wait_for_timeout(TEMPO_AGUARDAR_APOS_CNPJ_MS)

    inicio = time.time()
    timeout_s = TIMEOUT_CARREGAMENTO_COTAS_MS / 1000
    while (time.time() - inicio) < timeout_s:
        if page.locator(SELETORES["linhas_cota_unica"]).count() > 0:
            return True, ""

        erro_debito = detectar_erro_debito(page)
        if erro_debito:
            return False, f"Erro no portal: {erro_debito}"

        page.wait_for_timeout(500)

    return False, f"Cotas nao carregaram em ate {int(timeout_s)}s"


def baixar_duas_cotas(page, pasta_empresa: str, nome_empresa: str) -> bool:
    linhas = page.locator(SELETORES["linhas_cota_unica"])
    total_linhas = linhas.count()

    if total_linhas == 0:
        print("AVISO: nenhuma linha de Cota Unica / 1 Parcela encontrada.")
        return False

    baixados = 0
    for i in range(total_linhas):
        linha = linhas.nth(i)
        botao_imprimir = linha.locator(SELETORES["botao_imprimir_linha"]).first
        indice = i + 1
        nome_arquivo = f"{limpar_nome_pasta(nome_empresa)}_cota_unica_{indice}.pdf"
        caminho_final = os.path.join(pasta_empresa, nome_arquivo)

        try:
            with page.expect_download(timeout=TIMEOUT_PADRAO_MS) as download_info:
                botao_imprimir.click()
            download = download_info.value
            download.save_as(caminho_final)
            baixados += 1
            print(f"OK: download salvo -> {caminho_final}")
        except PlaywrightTimeoutError:
            print(f"AVISO: cota {indice} nao baixada (timeout/click).")
            continue

    if baixados < 2:
        print(f"AVISO: esperado 2 downloads, obtido {baixados}.")
        return False
    return True


def processar_empresa(page, empresa: Empresa) -> bool:
    print(f"Processando: {empresa.nome} | CNPJ: {empresa.cnpj}")
    pasta_empresa = preparar_pasta_empresa(empresa.nome)
    carregou, motivo_erro = preencher_cnpj_e_pesquisar(page, empresa.cnpj)
    if not carregou:
        motivo = motivo_erro or "Falha ao carregar cotas"
        print(f"AVISO: empresa pulada nesta tentativa: {empresa.nome}. Motivo: {motivo}")
        salvar_print_erro(page, pasta_empresa)
        registrar_erro_empresa(empresa, motivo)
        return False
    sucesso = baixar_duas_cotas(page, pasta_empresa, empresa.nome)
    if not sucesso:
        salvar_print_erro(page, pasta_empresa)
        registrar_erro_empresa(empresa, "Falha ao baixar as duas cotas")
    time.sleep(TEMPO_ENTRE_EMPRESAS_S)
    return sucesso


def executar_robo(
    empresas: list[Empresa],
    linhas_csv: list[list[str]],
    idx_coluna_mes: int,
    encoding_csv: str,
) -> None:
    with sync_playwright() as p:
        for empresa in empresas:
            browser = None
            context = None
            page = None
            try:
                browser = p.chromium.launch(headless=HEADLESS)
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()
                abrir_site(page)
                sucesso = processar_empresa(page, empresa)
            except Exception as exc:
                sucesso = False
                print(f"ERRO: falha no processamento da empresa {empresa.nome}: {exc}")
                if page is not None:
                    pasta_empresa = preparar_pasta_empresa(empresa.nome)
                    salvar_print_erro(page, pasta_empresa)
                registrar_erro_empresa(empresa, f"Excecao no processamento: {exc}")
            finally:
                if context:
                    context.close()
                if browser:
                    browser.close()

            if sucesso:
                registrar_mes_processado(
                    linhas_csv=linhas_csv,
                    idx_coluna_mes=idx_coluna_mes,
                    empresa=empresa,
                    valor=VALOR_MARCADOR_MES,
                )
                salvar_csv(CAMINHO_LISTA_EMPRESAS, linhas_csv, encoding_csv)


def main() -> None:
    print("Iniciando baixador de cota unica (esqueleto bruto).")
    print(f"Lista esperada: {CAMINHO_LISTA_EMPRESAS}")
    print(f"Pasta de saida: {PASTA_SAIDA}")
    empresas, linhas_csv, _, encoding_csv = carregar_empresas(CAMINHO_LISTA_EMPRESAS)
    coluna_mes = nome_coluna_mes(datetime.now())
    idx_coluna_mes = garantir_coluna_mes(linhas_csv, coluna_mes)
    executar_robo(empresas, linhas_csv, idx_coluna_mes, encoding_csv)


if __name__ == "__main__":
    main()
