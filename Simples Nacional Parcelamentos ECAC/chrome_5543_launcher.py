import os
import winreg
import undetected_chromedriver as uc
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PROFILE_ID = os.getenv("SCRIPT_PROFILE_ID", "5543").strip()
ECAC_PROFILE_PATH = os.getenv(
    "ECAC_PROFILE_PATH",
    os.path.join(BASE_DIR, f"chrome_perfil_{SCRIPT_PROFILE_ID}"),
)
CHROME_VERSION_MAIN = os.getenv("CHROME_VERSION_MAIN", "").strip()


def detectar_major_chrome():
    if CHROME_VERSION_MAIN.isdigit():
        return int(CHROME_VERSION_MAIN)
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon") as chave:
            versao, _ = winreg.QueryValueEx(chave, "version")
        major = str(versao).split(".", 1)[0]
        if major.isdigit():
            return int(major)
    except OSError:
        pass
    return None


def configurar_opcoes():
    options = uc.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--lang=pt-BR")
    options.add_argument("--disable-features=IsolateOrigins,site-per-process")
    os.makedirs(ECAC_PROFILE_PATH, exist_ok=True)
    options.add_argument(f"--user-data-dir={ECAC_PROFILE_PATH}")
    options.add_argument("--profile-directory=Default")
    return options


def ocultar_automacao(driver):
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        """
    })


def abrir_chrome():
    options = configurar_opcoes()
    major = detectar_major_chrome()
    if major is not None:
        print(f"✅ Chrome detectado (major): {major}")
        driver = uc.Chrome(options=options, version_main=major, use_subprocess=True)
    else:
        print("⚠️ Versão do Chrome não detectada, usando padrão.")
        driver = uc.Chrome(options=options, use_subprocess=True)

    driver.set_page_load_timeout(30)
    ocultar_automacao(driver)
    print(f"✅ Perfil: {ECAC_PROFILE_PATH}")
    return driver


if __name__ == "__main__":
    driver = None
    try:
        driver = abrir_chrome()
        driver.get("about:blank")
        input("Pressione Enter para fechar...")
    finally:
        if driver:
            driver.quit()