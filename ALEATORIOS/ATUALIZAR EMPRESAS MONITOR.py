import time
from playwright.sync_api import sync_playwright

def run(playwright):
    browser = playwright.chromium.launch(headless=False)  # coloque headless=True se quiser oculto
    page = browser.new_page()

    # --- LOGIN ---
    page.goto("https://app.monitorcontabil.com.br/login")

    page.get_by_role("textbox", name="E-mail").fill("contabil20@wilsonlopes.com.br")
    page.get_by_role("textbox", name="E-mail").press("Tab")
    page.get_by_role("textbox", name="*******").fill("email2016$")
    page.get_by_role("button", name="Entrar").click()

    time.sleep(1)

    page.goto("https://app.monitorcontabil.com.br/usuario/visualizar?busca=")

    time.sleep(1)

    # Aguarda carregar lista de usuários
    page.wait_for_selector(".lista-grid")

    # Captura lista de linhas
    user_rows = page.locator(".lista-grid table tbody tr")
    total_users = user_rows.count()

    print(f"Total de usuários encontrados: {total_users}")

    for i in range(total_users):
        
        # Ignorar o primeiro índice da lista (index 0)
        if i == 0:
            print("Ignorando o usuário no índice 0")
            continue

        row = user_rows.nth(i)

        # Captura o nome do usuário
        user_name = row.locator("td").first.inner_text().strip()

        print(f"Processando usuário: {user_name}")

        # Clica no botão da linha do usuário
        row.get_by_role("button").first.click()

        # Botão: Vincular Todas
        page.get_by_role("button", name="Vincular Todas", exact=True).click()

        # Botão: Salvar
        page.get_by_role("button", name="Salvar").click()

        # Aguarda voltar à lista
        page.wait_for_selector(".lista-grid")

    print("Processo finalizado!")
    browser.close()


with sync_playwright() as playwright:
    run(playwright)