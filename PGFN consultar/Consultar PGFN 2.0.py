import argparse
import csv
import json
import os
import re
import time
from datetime import datetime
from typing import Iterable, List, Optional, Tuple

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

URL_CONSULTA = "https://sisparnet.pgfn.fazenda.gov.br/sisparInternet/internet/darf/consultaParcelamentoDarfInternet.xhtml"
BASE_DIR = os.path.dirname(__file__)
ARQUIVO_JSON_DEFAULT = os.path.normpath(
    os.path.join(BASE_DIR, "..", "..", "central-utils", "data", "parcelamentos", "parcelamentos.import.json")
)
ARQUIVO_RESULTADO_CSV_DEFAULT = os.path.join(BASE_DIR, "LISTA PARCELAMENTOS - resultado.csv")
PASTA_SCREENSHOTS_DEFAULT = os.path.join(BASE_DIR, "screenshots")


def somente_numeros(valor: str) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def validar_documento(valor: str) -> str:
    doc = somente_numeros(valor)
    if len(doc) not in (11, 14):
        raise ValueError("CPF/CNPJ invalido. Informe 11 ou 14 digitos.")
    return doc


def validar_numero_parcelamento(valor: str) -> str:
    numero = somente_numeros(valor)
    if not numero:
        raise ValueError("Numero de parcelamento invalido.")
    return numero


def carregar_parcelamentos_json(caminho_json: str) -> List[Tuple[str, str, str]]:
    with open(caminho_json, "r", encoding="utf-8") as arquivo:
        payload = json.load(arquivo)

    itens = payload.get("items", []) if isinstance(payload, dict) else []
    dados: List[Tuple[str, str, str]] = []

    for item in itens:
        if not isinstance(item, dict):
            continue

        tipo = str(item.get("parcelamentoType") or "").strip().upper()
        if tipo != "PGFN":
            continue

        nome = str(item.get("companyName") or "").strip()
        cnpj_cpf = somente_numeros(item.get("cnpj"))
        numero_parcelamento = somente_numeros(item.get("parcelamentoNumber"))

        if len(cnpj_cpf) in (11, 14) and numero_parcelamento:
            dados.append((nome, cnpj_cpf, numero_parcelamento))

    return dados


def nome_arquivo_seguro(texto: str) -> str:
    texto = str(texto or "").strip()
    if not texto:
        return "SEM_NOME"
    return re.sub(r'[\\/:*?"<>|]+', "", texto).strip() or "SEM_NOME"


def salvar_screenshot_linha(
    driver: webdriver.Chrome,
    linha,
    pasta_screenshots: str,
    nome_dono: str,
) -> str:
    os.makedirs(pasta_screenshots, exist_ok=True)
    agora = datetime.now()
    dia_mes = agora.strftime("%d-%m")
    nome_arquivo = f"{nome_arquivo_seguro(nome_dono)} {dia_mes}.png"
    caminho = os.path.join(pasta_screenshots, nome_arquivo)

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", linha)
    time.sleep(0.4)
    driver.save_screenshot(caminho)
    return caminho


def detectar_mes_atual_status(
    driver: webdriver.Chrome,
    nome: str,
    cnpj_cpf: str,
    numero_parcelamento: str,
    pasta_screenshots: str,
) -> Tuple[bool, str, str, Optional[str]]:
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year

    tbody = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "formResumoParcelamentoDarf:tabaplicacaoparcelamentoparcela_data"))
    )
    linhas = tbody.find_elements(By.XPATH, "./tr")

    for linha in linhas:
        colunas = linha.find_elements(By.TAG_NAME, "td")
        if len(colunas) < 7:
            continue

        data_vencimento = colunas[3].text.strip()
        if not re.match(r"^\d{2}/\d{2}/\d{4}$", data_vencimento):
            continue

        _, mes, ano = data_vencimento.split("/")
        if int(mes) == mes_atual and int(ano) == ano_atual:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", linha)
            time.sleep(0.3)

            botoes_emissao = linha.find_elements(By.XPATH, ".//a[contains(@id,'NNimpressaoDarf')]")
            if any(botao.is_displayed() for botao in botoes_emissao):
                return True, "EMISSAO_DISPONIVEL", data_vencimento, None

            img_nao_emitido = linha.find_elements(
                By.XPATH,
                ".//img[contains(@id,'imprimirDarf') and "
                "(contains(translate(@title,'ÁÀÂÃÇÉÊÍÓÔÕÚáàâãçéêíóôõú','AAAACEEIOOOUaaaaceeiooou'),'Nao emitido') "
                "or contains(@title,'Não emitido'))]",
            )
            if any(img.is_displayed() for img in img_nao_emitido):
                return True, "EMISSAO_DISPONIVEL", data_vencimento, None

            img_ja_gerada = linha.find_elements(
                By.XPATH,
                ".//img["
                "contains(translate(@title,'ÁÀÂÃÇÉÊÍÓÔÕÚáàâãçéêíóôõú','AAAACEEIOOOUaaaaceeiooou'),'Emitido') "
                "or contains(translate(@title,'ÁÀÂÃÇÉÊÍÓÔÕÚáàâãçéêíóôõú','AAAACEEIOOOUaaaaceeiooou'),'Ja emitido') "
                "or contains(translate(@src,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'check') "
                "or contains(translate(@src,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'ok') "
                "or contains(translate(@src,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'sucesso') "
                "or contains(translate(@src,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'emitido')"
                "]",
            )

            botoes_alternativos = linha.find_elements(
                By.XPATH,
                ".//td[last()]//*[self::a or self::button or self::img]",
            )

            if any(img.is_displayed() for img in img_ja_gerada) or any(
                el.is_displayed() for el in botoes_alternativos
            ):
                screenshot = salvar_screenshot_linha(
                    driver=driver,
                    linha=linha,
                    pasta_screenshots=pasta_screenshots,
                    nome_dono=nome,
                )
                return True, "JA_GERADA", data_vencimento, screenshot

            return True, "SEM_BOTAO", data_vencimento, None

    return False, "NAO_ENCONTRADA", "", None


def consultar_uma_linha(
    driver: webdriver.Chrome,
    nome: str,
    cnpj_cpf: str,
    numero_parcelamento: str,
    pasta_screenshots: str,
) -> List[str]:
    driver.get(URL_CONSULTA)
    time.sleep(4)

    driver.find_element(By.ID, "formConsultaParcelamentoDarf:imNrCpfCnpjParcelamento").send_keys(cnpj_cpf)
    time.sleep(1)

    driver.switch_to.active_element.send_keys(Keys.TAB)
    time.sleep(1)

    driver.switch_to.active_element.send_keys(numero_parcelamento)
    time.sleep(1)

    iframe = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//iframe[contains(@title, 'hCaptcha')]"))
    )
    driver.switch_to.frame(iframe)
    driver.execute_script("document.querySelector('[role=\"checkbox\"]').click();")
    driver.switch_to.default_content()
    time.sleep(3)

    driver.find_element(By.ID, "formConsultaParcelamentoDarf:btnConsultar").click()
    time.sleep(3)

    encontrou_linha_mes, status_linha, data_linha, screenshot = detectar_mes_atual_status(
        driver=driver,
        nome=nome,
        cnpj_cpf=cnpj_cpf,
        numero_parcelamento=numero_parcelamento,
        pasta_screenshots=pasta_screenshots,
    )
    print(f"CPF/CNPJ: {cnpj_cpf} | Parc: {numero_parcelamento}")

    if encontrou_linha_mes and status_linha == "EMISSAO_DISPONIVEL":
        aviso = "Encontrado"
        print(f"OK: encontrou linha do mes atual ({data_linha}) com botao de emissao.")
    elif encontrou_linha_mes and status_linha == "JA_GERADA":
        aviso = "Ja gerada"
        print(f"OK: parcela do mes atual ({data_linha}) ja foi gerada (botao alternativo).")
        if screenshot:
            print(f"Screenshot salvo em: {screenshot}")
    elif encontrou_linha_mes and status_linha == "SEM_BOTAO":
        aviso = "Encontrado sem botao"
        print(f"ATENCAO: encontrou linha do mes atual ({data_linha}) sem botao de emissao.")
    else:
        aviso = "Nao encontrado"
        print("ATENCAO: nao encontrou linha do mes atual na tabela.")

    return [nome, cnpj_cpf, numero_parcelamento, aviso]


def escrever_csv(caminho_csv: str, linhas: Iterable[List[str]]) -> None:
    with open(caminho_csv, "w", newline="", encoding="utf-8-sig") as arquivo_csv:
        escritor = csv.writer(arquivo_csv, delimiter=";")
        escritor.writerow(["A", "B", "C", "D"])
        escritor.writerows(linhas)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Consulta status de parcelamento PGFN por CPF/CNPJ e numero de parcelamento."
    )
    parser.add_argument("--cnpj-cpf", help="CPF/CNPJ unico para consulta")
    parser.add_argument("--numero-parcelamento", help="Numero do parcelamento para consulta unica")
    parser.add_argument("--nome", default="", help="Nome para registrar no CSV da consulta unica")
    parser.add_argument("--arquivo-json", default=ARQUIVO_JSON_DEFAULT, help="JSON de entrada para modo lote")
    parser.add_argument("--resultado-csv", default=ARQUIVO_RESULTADO_CSV_DEFAULT, help="CSV de saida")
    parser.add_argument(
        "--pasta-screenshots",
        default=PASTA_SCREENSHOTS_DEFAULT,
        help="Pasta para screenshots de casos ja gerados",
    )
    parser.add_argument("--manter-aberto", action="store_true", help="Nao fecha o navegador ao final")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    consulta_unica = bool(args.cnpj_cpf or args.numero_parcelamento)
    if consulta_unica and not (args.cnpj_cpf and args.numero_parcelamento):
        print("ERRO: para consulta unica informe --cnpj-cpf e --numero-parcelamento juntos.")
        return 2

    try:
        if consulta_unica:
            dados = [
                (
                    str(args.nome or "").strip(),
                    validar_documento(args.cnpj_cpf),
                    validar_numero_parcelamento(args.numero_parcelamento),
                )
            ]
        else:
            if not os.path.exists(args.arquivo_json):
                print(f"ERRO: arquivo nao encontrado: {args.arquivo_json}")
                return 2
            dados = carregar_parcelamentos_json(args.arquivo_json)
            if not dados:
                raise ValueError("Nenhum registro PGFN valido encontrado no JSON.")
    except ValueError as e:
        print(f"ERRO: {e}")
        return 2

    driver = webdriver.Chrome()
    resultados_csv: List[List[str]] = []

    try:
        for nome, cnpj_cpf, numero_parcelamento in dados:
            try:
                resultados_csv.append(
                    consultar_uma_linha(
                        driver=driver,
                        nome=nome,
                        cnpj_cpf=cnpj_cpf,
                        numero_parcelamento=numero_parcelamento,
                        pasta_screenshots=args.pasta_screenshots,
                    )
                )
            except Exception as e:
                print(f"Falha no registro {cnpj_cpf}/{numero_parcelamento}: {e}")
                resultados_csv.append([nome, cnpj_cpf, numero_parcelamento, f"Erro: {e}"])

        escrever_csv(args.resultado_csv, resultados_csv)
        print(f"CSV salvo em: {args.resultado_csv}")
        return 0
    finally:
        if args.manter_aberto:
            print("Navegador mantido aberto. Feche manualmente quando terminar.")
            try:
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                pass
        driver.quit()


if __name__ == "__main__":
    raise SystemExit(main())
