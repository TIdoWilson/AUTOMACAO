import json
import os
import time
import zipfile
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver as RemoteWebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

# ================== CONFIGURAÇÕES ==================

DIR_ORIGEM = r"C:\Users\Usuario\Desktop\teste"

# mesmo arquivo de sessão que o 01_abre_iob_firefox.py cria
ARQ_SESSAO = os.path.join(
    os.path.dirname(__file__),
    "iob_firefox_session.json"
)

# Seletores
SEL_BTN_ACESSAR_PAINEL = (
    By.CSS_SELECTOR,
    "button[data-testid='link-status-sped']"
)

SEL_BTN_IMPORTAR_SPED = (
    By.XPATH,
    "//button[contains(., 'Importar Arquivo SPED')]"
)

SEL_SPAN_UPLOAD = (
    By.XPATH,
    "//span[contains(., 'upload')]"
)

SEL_BTN_CONFIRMAR = (
    By.XPATH,
    "//button[contains(., 'Confirmar')]"
)

# ================== FUNÇÕES AUXILIARES ==================


def criar_zip_com_efds(dir_origem: str) -> str:
    """Cria um ZIP com os arquivos da pasta dir_origem (recursivo),
    ignorando qualquer arquivo .zip já existente.
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = os.path.join(dir_origem, f"EFDs_{ts}.zip")

    print(f"Criando ZIP: {zip_path}")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(dir_origem):
            for fname in files:
                # ignora todos os .zip
                if fname.lower().endswith(".zip"):
                    continue

                fpath = os.path.join(root, fname)
                rel = os.path.relpath(fpath, dir_origem)
                print(f"  adicionando: {fpath} -> {rel}")
                zf.write(fpath, rel)

    print("ZIP criado.")
    return zip_path


def carregar_sessao_salva(caminho: str):
    if not os.path.exists(caminho):
        print(f"Arquivo de sessão não encontrado: {caminho}")
        return None

    with open(caminho, "r", encoding="utf-8") as f:
        data = json.load(f)

    executor_url = data.get("executor_url")
    session_id = data.get("session_id")

    if not executor_url or not session_id:
        print("Arquivo de sessão não contém executor_url ou session_id.")
        return None

    return executor_url, session_id


def conectar_firefox_existente():
    """Tenta reutilizar APENAS a sessão salva.
       Se não conseguir, levanta erro e NÃO abre novo Firefox.
    """
    sessao = carregar_sessao_salva(ARQ_SESSAO)
    if sessao is None:
        raise RuntimeError(
            "Sessão do Firefox não encontrada. "
            "Execute primeiro o script 01_abre_iob_firefox.py e depois rode este."
        )

    executor_url, session_id = sessao
    print("Reconectando ao Firefox existente...")
    print(f"  executor_url = {executor_url}")
    print(f"  session_id   = {session_id}")

    try:
        from selenium.webdriver.firefox.options import Options as FirefoxOptions

        opts = FirefoxOptions()
        driver = RemoteWebDriver(
            command_executor=executor_url,
            options=opts
        )
        driver.session_id = session_id

        # teste rápido para verificar se a sessão está viva
        _ = driver.title
        print("Conexão com Firefox reutilizado estabelecida.")
        return driver

    except Exception as e:
        print("Falha ao reconectar à sessão Firefox salva.")
        print(f"Detalhe técnico: {e}")
        raise RuntimeError(
            "Não foi possível reutilizar a sessão existente do Firefox.\n"
            "Certifique-se de que o Firefox (e o geckodriver) ainda estão abertos, "
            "e execute novamente o script 01_abre_iob_firefox.py antes de rodar o Upload_Sped_IOB.py."
        ) from e


def acessar_painel_e_importar_sped(driver, zip_path: str):
    wait = WebDriverWait(driver, 60)

    try:
        # 1) Clicar em "Acessar painel"
        print("Clicando em 'Acessar painel'...")
        btn_painel = wait.until(EC.element_to_be_clickable(SEL_BTN_ACESSAR_PAINEL))
        btn_painel.click()

        time.sleep(15)  # aguarda painel carregar

        # 2) Clicar em "Importar Arquivo SPED"
        print("Clicando em 'Importar Arquivo SPED' (1ª vez)...")
        btn_import1 = wait.until(EC.element_to_be_clickable(SEL_BTN_IMPORTAR_SPED))
        btn_import1.click()

        time.sleep(15)

        # 3) Clicar em "Importar Arquivo SPED" (segunda vez)
        print("Clicando em 'Importar Arquivo SPED' (2ª vez)...")
        btn_import2 = wait.until(EC.element_to_be_clickable(SEL_BTN_IMPORTAR_SPED))
        btn_import2.click()

        time.sleep(15)

        # 4) No modal, clicar em "upload"
        print("Clicando em 'upload' no modal...")
        span_upload = wait.until(EC.element_to_be_clickable(SEL_SPAN_UPLOAD))
        span_upload.click()

        time.sleep(2)

        # 5) Enviar caminho do arquivo ZIP via <input type="file">
        print("Procurando campo <input type='file'> para enviar o ZIP...")
        input_file = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
        input_file.send_keys(zip_path)

        time.sleep(10)  # aguarda upload

        # 6) Clicar no botão "Confirmar"
        print("Clicando no botão 'Confirmar'...")
        btn_confirmar = wait.until(EC.element_to_be_clickable(SEL_BTN_CONFIRMAR))
        btn_confirmar.click()

        print("Upload enviado. Aguardando processamento no site...")
        time.sleep(15)

    except WebDriverException as e:
        print("\nERRO DE COMUNICAÇÃO COM O FIREFOX / SELENIUM.")
        print("A sessão provavelmente foi perdida (geckodriver caiu ou porta fechou).")
        print(f"Detalhe técnico: {e}")
        print(
            "\nAção necessária:\n"
            "- Feche o Firefox se estiver inconsistente.\n"
            "- Rode novamente o 01_abre_iob_firefox.py, faça o login manual.\n"
            "- Depois execute de novo o Upload_Sped_IOB.py."
        )
        return


# ================== MAIN ==================


def main():
    # 1) Conecta SOMENTE à sessão Firefox existente
    try:
        driver = conectar_firefox_existente()
    except RuntimeError as e:
        print("\nERRO CRÍTICO DE SESSÃO SELENIUM:")
        print(e)
        return

    # 2) Cria ZIP com as EFDs (só se a sessão está viva)
    zip_path = criar_zip_com_efds(DIR_ORIGEM)

    # 3) Faz o fluxo de upload usando o Firefox já logado
    acessar_painel_e_importar_sped(driver, zip_path)
    print("\nFim do script Upload_Sped_IOB.py (do ponto de vista do script).")


if __name__ == "__main__":
    main()
