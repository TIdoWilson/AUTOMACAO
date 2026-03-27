import json
import os
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options

from env_loader import load_local_env


BASE_DIR = Path(__file__).resolve().parent
SESSION_FILE = BASE_DIR / "iob_firefox_session.json"
load_local_env(str(BASE_DIR / ".env"))

IOB_LOGIN_URL = os.getenv(
    "IOB_LOGIN_URL",
    (
        "https://sso.iob.com.br/signin/?response_type=code"
        "&scope=&client_id=b89f6d4c-78bb-4995-9ac7-aa5459f5cf6d"
        "&redirect_uri=https%3A%2F%2Fapp.iob.com.br%2Fcallback%2F%3Fpath%3Dhttps%3A%2F%2Fapp.iob.com.br%2Fapp%2F"
        "&isSignUpDisable=false&showFAQ=false&isSocialLoginDisable=false"
    ),
)

# porta fixa para o geckodriver (usada depois para reconectar)
GECKO_PORT = 4444


def abrir_firefox_e_logar():
    # geckodriver precisa estar no PATH
    service = Service(port=GECKO_PORT)

    options = Options()
    options.add_argument("--width=1400")
    options.add_argument("--height=900")

    driver = webdriver.Firefox(service=service, options=options)
    driver.get(IOB_LOGIN_URL)

    print("\nFirefox aberto na tela de login da IOB.")
    print("1) Faça o login manualmente no navegador (SSO).")
    print("2) Aguarde carregar o app/painel da IOB (onde aparece 'Acessar painel').")
    input("3) Quando o login estiver concluído e o painel carregado, pressione ENTER aqui no terminal... ")

    # dados necessários para reconectar depois
    session_data = {
        "session_id": driver.session_id,
        # como fixamos a porta do geckodriver, o executor_url é conhecido
        "executor_url": f"http://127.0.0.1:{GECKO_PORT}",
    }

    with open(SESSION_FILE, "w", encoding="utf-8") as f:
        json.dump(session_data, f)

    print(f"\nSessão salva em: {SESSION_FILE}")
    print("IMPORTANTE: NÃO feche o Firefox nem a aba do sistema IOB.")
    print("Este script vai encerrar, mas o navegador continuará aberto e será reutilizado pelo script de upload.")

    # não chamar driver.quit(): queremos manter o Firefox vivo
    time.sleep(1)


def main():
    abrir_firefox_e_logar()


if __name__ == "__main__":
    main()
