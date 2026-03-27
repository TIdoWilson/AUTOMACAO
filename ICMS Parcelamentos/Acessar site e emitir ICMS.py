import argparse
import os
import re
import time
from calendar import monthrange
from datetime import date

import requests
from selenium import webdriver
from selenium.common.exceptions import NoSuchWindowException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def somente_numeros(valor: str) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def validar_cnpj(cnpj: str) -> str:
    cnpj_limpo = somente_numeros(cnpj)
    if len(cnpj_limpo) != 14:
        raise ValueError("CNPJ invalido. Informe 14 digitos.")
    return cnpj_limpo


def validar_termo(termo: str) -> str:
    termo_limpo = somente_numeros(termo)
    if not termo_limpo:
        raise ValueError("Numero/termo de parcelamento invalido.")
    return termo_limpo


def criar_opcoes_chrome(download_folder: str) -> webdriver.ChromeOptions:
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    return chrome_options


def log_erro(script_dir: str, cnpj: str, termo_parcelamento: str, mensagem: str) -> None:
    with open(os.path.join(script_dir, "erro_emissao_guia.txt"), "a", encoding="utf-8") as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {mensagem}\n")
        f.write(f"   CNPJ: {cnpj}\n")
        f.write(f"   Termo: {termo_parcelamento}\n")
        f.write("-" * 50 + "\n")


def baixar_pdf(url: str, nome_arquivo: str, download_folder: str) -> bool:
    try:
        response = requests.get(url, stream=True, timeout=30)
        if response.status_code == 200 and "application/pdf" in response.headers.get("Content-Type", ""):
            caminho = os.path.join(download_folder, nome_arquivo)
            with open(caminho, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"PDF baixado: {caminho}")
            return True

        print(
            "Resposta nao e PDF: "
            f"status={response.status_code}, content-type={response.headers.get('Content-Type')}"
        )
        return False
    except Exception as e:
        print(f"Erro no download: {e}")
        return False


def executar_fluxo(cnpj: str, termo_parcelamento: str, pausar_final: bool = False) -> int:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    download_folder = os.path.join(script_dir, "PDFs Requisitados")
    os.makedirs(download_folder, exist_ok=True)

    hoje = date.today()
    ultimo_dia = monthrange(hoje.year, hoje.month)[1]
    data_pagamento = f"{ultimo_dia:02d}/{hoje.month:02d}/{hoje.year}"

    chrome_options = criar_opcoes_chrome(download_folder)
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    driver.get("https://emitirgrpr.sefa.pr.gov.br/arrecadacao/emitir/guiatela")
    wait = WebDriverWait(driver, 15)

    try:
        try:
            cookie_btn = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(text(), 'Rejeitar') or contains(text(), 'Recusar')]")
                )
            )
            cookie_btn.click()
            print("Cookies recusados")
            time.sleep(1)
        except TimeoutException:
            print("Nenhum banner de cookies encontrado")

        container_cat = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class, 'multiselect') and .//input[@id='codCategoria']]")
            )
        )
        container_cat.click()
        time.sleep(0.5)

        campo_cat = wait.until(EC.presence_of_element_located((By.ID, "codCategoria")))
        campo_cat.clear()
        campo_cat.send_keys("PARCELAMENTO")
        time.sleep(1)

        opcao_parc = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'PARCELAMENTO')]")))
        opcao_parc.click()

        btn_avancar1 = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-primary') and contains(text(), 'Avançar')]"))
        )
        btn_avancar1.click()

        wait.until(EC.presence_of_element_located((By.ID, "id_0_2")))
        time.sleep(2)

        tipo_container = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class, 'multiselect') and .//span[contains(text(), 'CNPJ')]]")
            )
        )
        tipo_container.click()
        time.sleep(0.5)

        opcao_cnpj = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'CNPJ')]")))
        opcao_cnpj.click()

        campo_cnpj = wait.until(EC.presence_of_element_located((By.ID, "id_0_2")))
        campo_cnpj.clear()
        campo_cnpj.send_keys(cnpj)

        campo_termo = wait.until(EC.presence_of_element_located((By.ID, "id_0_3")))
        campo_termo.clear()
        campo_termo.send_keys(termo_parcelamento)

        parcela_container = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class, 'multiselect') and .//span[contains(text(), 'Próxima')]]")
            )
        )
        parcela_container.click()
        time.sleep(0.5)
        opcao_proxima = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Próxima')]")))
        opcao_proxima.click()

        data_container = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'multiselect') and .//input[contains(@id, 'id_0_6')]]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", data_container)
        time.sleep(0.5)
        data_container.click()
        time.sleep(1)

        opcao_data = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()='{data_pagamento}']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", opcao_data)
        time.sleep(0.5)
        opcao_data.click()

        btn_avancar2 = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-primary') and contains(text(), 'Avançar')]"))
        )
        btn_avancar2.click()

        try:
            botao_emitir = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-success') and contains(text(), 'Emitir Guia')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_emitir)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", botao_emitir)
        except TimeoutException:
            erro = "Botao 'Emitir Guia' nao apareceu"
            print(erro)
            log_erro(script_dir, cnpj, termo_parcelamento, erro)
            return 1

        print("Aguardando carregamento apos 'Emitir Guia'...")
        time.sleep(5)

        def clicar_salvar_pdf() -> bool:
            try:
                botao = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Salvar PDF')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", botao)
                return True
            except Exception:
                pass

            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            for iframe in iframes:
                driver.switch_to.frame(iframe)
                try:
                    botao = driver.find_element(By.XPATH, "//*[contains(text(), 'Salvar PDF')]")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", botao)
                    driver.switch_to.default_content()
                    return True
                except Exception:
                    driver.switch_to.default_content()
            return False

        start = time.time()
        salvar_clicado = False
        while time.time() - start < 20:
            if clicar_salvar_pdf():
                salvar_clicado = True
                break
            time.sleep(1)

        if not salvar_clicado:
            html_path = os.path.join(script_dir, "pagina_erro_salvar.html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            erro = f"Botao 'Salvar PDF' nao encontrado. Pagina salva em: {html_path}"
            print(erro)
            log_erro(script_dir, cnpj, termo_parcelamento, erro)
            return 1

        original_window = driver.current_window_handle
        print("Aguardando abertura da nova janela...")
        try:
            WebDriverWait(driver, 20).until(EC.number_of_windows_to_be(2))
            for handle in driver.window_handles:
                if handle != original_window:
                    driver.switch_to.window(handle)
                    break

            time.sleep(3)
            html_nova = os.path.join(script_dir, "nova_janela.html")
            with open(html_nova, "w", encoding="utf-8") as f:
                f.write(driver.page_source)

            url_pdf = driver.current_url
            nome_arquivo = f"GR_{cnpj}_{termo_parcelamento}_{data_pagamento.replace('/', '-')}.pdf"
            if not baixar_pdf(url_pdf, nome_arquivo, download_folder):
                log_erro(script_dir, cnpj, termo_parcelamento, f"Falha no download do PDF a partir de {url_pdf}")
                return 1

            driver.close()
            driver.switch_to.window(original_window)

        except TimeoutException:
            print("Nenhuma nova janela detectada. Tentando baixar PDF da pagina atual...")
            url_pdf = driver.current_url
            if "pdf" in url_pdf.lower() or driver.page_source.startswith("%PDF"):
                nome_arquivo = f"GR_{cnpj}_{termo_parcelamento}_{data_pagamento.replace('/', '-')}.pdf"
                if not baixar_pdf(url_pdf, nome_arquivo, download_folder):
                    log_erro(script_dir, cnpj, termo_parcelamento, "Falha ao baixar PDF da pagina atual")
                    return 1
            else:
                html_path = os.path.join(script_dir, "pagina_sem_pdf.html")
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                log_erro(script_dir, cnpj, termo_parcelamento, f"PDF nao localizado. Pagina salva em: {html_path}")
                return 1

        return 0

    except NoSuchWindowException as e:
        log_erro(script_dir, cnpj, termo_parcelamento, f"Janela fechada inesperadamente: {e}")
        return 1
    except Exception as e:
        log_erro(script_dir, cnpj, termo_parcelamento, f"Erro geral: {e}")
        return 1
    finally:
        if pausar_final:
            input("Pressione Enter para fechar...")
        driver.quit()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Emite guia ICMS informando CNPJ e numero do parcelamento.")
    parser.add_argument("--cnpj", required=True, help="CNPJ da empresa (14 digitos)")
    parser.add_argument(
        "--numero-parcelamento",
        "--termo",
        dest="numero_parcelamento",
        required=True,
        help="Numero/termo do parcelamento ICMS",
    )
    parser.add_argument("--pausar-final", action="store_true", help="Aguarda Enter ao final")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        cnpj = validar_cnpj(args.cnpj)
        termo = validar_termo(args.numero_parcelamento)
    except ValueError as e:
        print(f"ERRO: {e}")
        return 2

    return executar_fluxo(cnpj, termo, pausar_final=args.pausar_final)


if __name__ == "__main__":
    raise SystemExit(main())
