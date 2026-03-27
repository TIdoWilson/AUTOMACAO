import os
import time
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- Configurações ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(BASE_DIR, "cnpjs.xlsx")
URL = "https://www3.mte.gov.br/sistemas/mediador/ConsultarInstColetivo"

# --- Função para criar driver com pasta de download personalizada ---
def get_chrome_driver(download_path):
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

# --- Ler CNPJs do Excel ---
print(f"📂 Lendo arquivo {EXCEL_PATH} ...")
df = pd.read_excel(EXCEL_PATH)
cnpjs = df.iloc[:, 0].dropna().astype(str).tolist()
print(f"✔ {len(cnpjs)} CNPJs carregados do Excel.")

# --- Função principal de download por CNPJ ---
def baixar_por_cnpj(cnpj):
    cnpj_dir = os.path.join(DOWNLOAD_DIR, cnpj)
    os.makedirs(cnpj_dir, exist_ok=True)

    driver = get_chrome_driver(cnpj_dir)
    wait = WebDriverWait(driver, 20)

    try:
        driver.get(URL)
        time.sleep(2)

        # --- Marcar checkbox CNPJ ---
        chk_cnpj = wait.until(EC.element_to_be_clickable((By.ID, "chkNRCNPJ")))
        driver.execute_script("arguments[0].click();", chk_cnpj)
        time.sleep(1)

        # --- Preencher CNPJ ---
        driver.execute_script(f"document.getElementById('txtNRCNPJ').value = '{cnpj}';")
        driver.execute_script("$('#txtNRCNPJ').trigger('change');")
        time.sleep(1)

        # --- Marcar checkbox Vigência e selecionar "Todos" ---
        chk_vig = wait.until(EC.element_to_be_clickable((By.ID, "chkVigencia")))
        driver.execute_script("arguments[0].click();", chk_vig)
        time.sleep(1)

        sel_vig = wait.until(EC.presence_of_element_located((By.ID, "cboSTVigencia")))
        driver.execute_script("arguments[0].value = '2';", sel_vig)
        driver.execute_script("$('#cboSTVigencia').trigger('change');")
        time.sleep(1)

        # --- Preencher período ---
        driver.execute_script("document.getElementById('txtDTInicioVigencia').value='01/01/2025';")
        driver.execute_script("document.getElementById('txtDTFimVigencia').value='31/12/2025';")
        time.sleep(1)

        # --- Clicar Pesquisar ---
        btn_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "btnPesquisar")))
        driver.execute_script("arguments[0].click();", btn_pesquisar)
        time.sleep(3)

        # --- Iterar páginas e baixar arquivos ---
        main_window = driver.current_window_handle
        while True:
            time.sleep(2)
            download_links = driver.find_elements(By.XPATH, "//a[contains(@onclick,'fDownload')]")

            for link in download_links:
                # Abrir link de download
                driver.execute_script("arguments[0].click();", link)
                time.sleep(2)  # espera o download iniciar

                # Alternar para a nova janela e fechar
                new_window = [w for w in driver.window_handles if w != main_window][0]
                driver.switch_to.window(new_window)
                time.sleep(2)  # tempo para o download começar
                driver.close()
                driver.switch_to.window(main_window)

            # Verificar próxima página
            next_buttons = driver.find_elements(By.XPATH, "//a[contains(text(), 'Próximo')]")
            if next_buttons and next_buttons[0].is_enabled():
                driver.execute_script("arguments[0].click();", next_buttons[0])
                time.sleep(3)
            else:
                break

        print(f"✅ Downloads do CNPJ {cnpj} concluídos.")

    except Exception as e:
        print(f"❌ Erro ao processar {cnpj}: {e}")
    finally:
        driver.quit()

# --- Loop principal ---
for cnpj in cnpjs:
    print(f"\n🔍 Processando CNPJ: {cnpj}")
    baixar_por_cnpj(cnpj)

print("\n🏁 Todos os downloads concluídos.")
