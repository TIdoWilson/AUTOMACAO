# econet_open_browser_persist_cdp.py
# Mantém o Chrome/Chromium aberto após o script encerrar usando CDP (remote debugging)

import os, re, time, subprocess, sys, shutil
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

USERNAME = os.environ.get("ECONET_USER", "ONU33962")
PASSWORD = os.environ.get("ECONET_PASS", "WilsonL2026")

LOGIN_URL = "https://www.econeteditora.com.br/user/login.asp"
LOGIN_SUCCESS_RE = re.compile(r"^https://www\.econeteditora\.com\.br/(novo/|inicial\.php)(\?.*)?$")
FINAL_LINK = "https://app.econeteditora.com.br/app/eco-class"
CLICK_TIMEOUT = 5
REMOTE_DEBUGGING_URL = "http://127.0.0.1:9222"
PROFILE_DIR = str(Path.home() / ".econet_chrome_profile")  # perfil persistente

# --- util: localizar Chrome/Chromium no sistema ---
def find_chrome_executable():
    candidates = []
    if sys.platform.startswith("win"):
        candidates += [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files\Chromium\Application\chrome.exe",
            r"C:\Program Files (x86)\Chromium\Application\chrome.exe",
            shutil.which("chrome"),
            shutil.which("chromium"),
            shutil.which("chromium-browser"),
        ]
    elif sys.platform == "darwin":
        candidates += [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
            shutil.which("google-chrome"),
            shutil.which("chromium"),
        ]
    else:  # Linux e afins
        candidates += [
            shutil.which("google-chrome"),
            shutil.which("google-chrome-stable"),
            shutil.which("chromium"),
            shutil.which("chromium-browser"),
        ]
    return next((c for c in candidates if c and Path(c).exists()), None)

def start_external_chrome(remote_port=9222, user_data_dir=PROFILE_DIR):
    chrome_path = find_chrome_executable()
    if not chrome_path:
        return None

    Path(user_data_dir).mkdir(parents=True, exist_ok=True)

    # Lança o Chrome de forma independente do Playwright (não bloqueia e não é fechado com o Python)
    args = [
        chrome_path,
        f"--remote-debugging-port={remote_port}",
        f"--user-data-dir={user_data_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
        "about:blank",
    ]

    # Inicia e não espera (o processo continuará independente do Python)
    creationflags = 0
    if sys.platform.startswith("win"):
        # Evita abrir janela de console extra
        creationflags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP

    try:
        subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=creationflags)
        return chrome_path
    except Exception as e:
        print(f"[ERRO] Falha ao iniciar Chrome: {e}")
        return None

def try_click_recaptcha_checkbox(page):
    try:
        for frame in page.frames:
            for sel in ['#recaptcha-anchor', '.recaptcha-checkbox-border', 'div.recaptcha-checkbox-checkmark']:
                el = frame.query_selector(sel)
                if el:
                    el.click(timeout=CLICK_TIMEOUT * 1000)
                    return True
    except Exception as e:
        print("Erro reCAPTCHA:", e)
    return False

def wait_for_manual_login(context, page):
    while True:
        if LOGIN_SUCCESS_RE.match(page.url):
            return page
        for p in context.pages:
            if LOGIN_SUCCESS_RE.match(p.url):
                return p
        time.sleep(1)

def main():
    # 1) Garante um Chrome externo rodando com remote debugging
    chrome_path = start_external_chrome(remote_port=9222, user_data_dir=PROFILE_DIR)
    if not chrome_path:
        print("\n[ATENÇÃO] Não foi possível iniciar o Chrome automaticamente.")
        print("Inicie manualmente e depois rode o script novamente, por exemplo:")
        print("  Windows:")
        print(r'    "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\.econet_chrome_profile"')
        print("  macOS:")
        print(r'    /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222 --user-data-dir="$HOME/.econet_chrome_profile"')
        print("  Linux:")
        print(r'    google-chrome --remote-debugging-port=9222 --user-data-dir="$HOME/.econet_chrome_profile"')
        return

    # 2) Conecta via CDP ao Chrome externo
    pw = sync_playwright().start()
    try:
        browser = pw.chromium.connect_over_cdp(REMOTE_DEBUGGING_URL)

        # Pega o primeiro contexto (padrão) ou cria uma nova página nele
        contexts = browser.contexts
        if contexts:
            context = contexts[0]
        else:
            # Normalmente sempre há um contexto padrão; esta linha é fallback
            context = browser.new_context()

        page = context.new_page()
        page.goto(LOGIN_URL)

        page.fill('input[name="Log"]', USERNAME)
        page.fill('input[name="Sen"]', PASSWORD)
        try_click_recaptcha_checkbox(page)

        btn = page.query_selector('input[type="submit"], button[type="submit"]')
        if btn:
            try:
                btn.click()
            except Exception:
                pass

        try:
            page.wait_for_url(LOGIN_SUCCESS_RE, timeout=60000)
        except PlaywrightTimeoutError:
            page = wait_for_manual_login(context, page)

        page.goto(FINAL_LINK)
        print("✅ Login concluído. Navegador externo conectado e na página final.")
        print("ℹ️ Você pode fechar este terminal/rodar 'python ...' até o fim — o Chrome continuará aberto.")

        # (Opcional) Salva estado (cookies/localStorage) no perfil persistente
        # Em CDP, o perfil já está no --user-data-dir, então o estado persiste automaticamente.

    finally:
        # 3) IMPORTANTÍSSIMO: apenas desligue o Playwright (desconecta do CDP).
        # NÃO fechamos browser/context/page — o Chrome é externo e continua aberto.
        pw.stop()

    print("Script finalizado. O Chrome/Chromium permanece aberto.")

if __name__ == "__main__":
    main()
