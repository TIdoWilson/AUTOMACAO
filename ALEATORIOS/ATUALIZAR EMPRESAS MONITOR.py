import os
import time
from pathlib import Path

from playwright.sync_api import sync_playwright


def load_local_env(*paths: Path) -> None:
    for path in paths:
        if not path.is_file():
            continue
        try:
            with path.open("r", encoding="utf-8") as handle:
                for line in handle:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()
                    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
                        value = value[1:-1]
                    if key and key not in os.environ:
                        os.environ[key] = value
        except OSError:
            continue


BASE_DIR = Path(__file__).resolve().parent
load_local_env(BASE_DIR / ".env")

EMAIL = os.getenv("IOB_USUARIO", "").strip()
PASSWORD = os.getenv("IOB_SENHA", "").strip()


def run(playwright):
    browser = playwright.chromium.launch(headless=False)
    page = browser.new_page()

    page.goto("https://app.monitorcontabil.com.br/login")

    if not EMAIL or not PASSWORD:
        raise RuntimeError("Defina IOB_USUARIO e IOB_SENHA em ALEATORIOS/.env")

    page.get_by_role("textbox", name="E-mail").fill(EMAIL)
    page.get_by_role("textbox", name="E-mail").press("Tab")
    page.get_by_role("textbox", name="*******").fill(PASSWORD)
    page.get_by_role("button", name="Entrar").click()

    time.sleep(1)

    page.goto("https://app.monitorcontabil.com.br/usuario/visualizar?busca=")
    time.sleep(1)

    page.wait_for_selector(".lista-grid")

    user_rows = page.locator(".lista-grid table tbody tr")
    total_users = user_rows.count()

    print(f"Total de usuarios encontrados: {total_users}")

    for i in range(total_users):
        if i == 0:
            print("Ignorando o usuario no indice 0")
            continue

        row = user_rows.nth(i)
        user_name = row.locator("td").first.inner_text().strip()

        print(f"Processando usuario: {user_name}")

        row.get_by_role("button").first.click()
        page.get_by_role("button", name="Vincular Todas", exact=True).click()
        page.get_by_role("button", name="Salvar").click()
        page.wait_for_selector(".lista-grid")

    print("Processo finalizado!")
    browser.close()


if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
