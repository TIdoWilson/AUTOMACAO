import os, re, time
from datetime import datetime
from pathlib import Path
import pandas as pd
from dateutil.relativedelta import relativedelta
from playwright.sync_api import sync_playwright, expect
from chrome_9222 import chrome_9222, PORT

CSV_NAME = "DOMESTICAS.csv"

def click_and_wait_dom(page, locator, timeout=60000):
    """Clica e espera 'domcontentloaded' se houver navegação; senão apenas garante estabilidade."""
    with page.expect_navigation(wait_until="domcontentloaded", timeout=timeout) as maybe_nav:
        locator.click()
    # se não houve navegação, o expect_navigation sai por timeout; tratamos isso
    try:
        maybe_nav.value
    except Exception:
        pass

def ensure_same_page_or_popup(context, click_locator, timeout=60000):
    """Alguns botões podem abrir popup/aba nova. Trata ambos os casos e retorna a 'page' adequada."""
    page = context.pages[0]
    with context.expect_page(timeout=timeout) as new_page_info:
        click_locator.click()
    try:
        new_page = new_page_info.value
        # Se abriu popup, use a nova página
        new_page.wait_for_load_state("domcontentloaded", timeout=timeout)
        return new_page
    except Exception:
        # Não abriu popup: continue na mesma page
        page.wait_for_load_state("domcontentloaded", timeout=timeout)
        return page

if __name__ == "__main__":
    with sync_playwright() as p:
        browser = chrome_9222(p, PORT)   # conecta ou inicia Chrome
        context = browser.contexts[0]    # contexto persistente
        if context.pages:
            page = context.pages[0]   # pega a primeira aba já aberta
        else:
            page = context.new_page()  # só cria se não houver nenhuma

        page.goto("https://login.esocial.gov.br/login.aspx")
        time.sleep(0.5)
        page.click("text=Entrar com")
        page.locator("button").filter(has_text=re.compile(r"Entrar com&nbsp", re.I))
        expect(page).to_have_url(re.compile(r".*/portal/?$"), timeout=60000)         
        print("acessou")

        
        # Seleciona o perfil "Procurador PF - CPF"
        select_perfil = page.locator("select#perfilAcesso")
        expect(select_perfil).to_be_visible(timeout=60000)
        select_perfil.select_option("PROCURADOR_PF")
        expect(select_perfil).to_have_value("PROCURADOR_PF")

        # Localizador do input do CPF (conforme seu print: input#procuradorCpf)
        cpf_input = page.locator("input#procuradorCpf.form-control.cpf.required, input#procuradorCpf")
        expect(cpf_input).to_be_visible(timeout=60000)

        base_dir = Path(__file__).resolve().parent
        csv_path = base_dir / CSV_NAME
        if not csv_path.exists():
            raise FileNotFoundError(f"Não achei {CSV_NAME} na mesma pasta do script: {csv_path}")

        # Tenta detectar o separador automaticamente; se falhar, tenta ; com latin1
        try:
            df = pd.read_csv(csv_path, dtype=str, sep=None, engine="python", encoding="utf-8")
        except Exception:
            df = pd.read_csv(csv_path, dtype=str, sep=";", encoding="latin1")

        # Descobre a coluna "CPF" (case-insensitive); se não existir, usa a 1ª coluna
        cpf_col = next((c for c in df.columns if c.strip().lower() == "cpf"), df.columns[0])

        cpfs = (
            df[cpf_col]
            .dropna()
            .astype(str)
            .str.replace(r"\D", "", regex=True)  # mantém só dígitos
            .str.zfill(11)                       # garante 11 dígitos
            .tolist()
        )

        for cpf in cpfs:
            # 1) Preencher CPF
            cpf_input.click()
            # limpa e digita (campo mascarado costuma aceitar fill; se não, use type)
            try:
                cpf_input.fill("")  # limpar
            except Exception:
                pass
            cpf_input.press("Control+a")
            cpf_input.type(cpf)
            # Validação rápida de que ficou algo no campo
            expect(cpf_input).to_have_value(re.compile(r"\d{3}.*"), timeout=5000)

            # 3) Clicar "Verificar"
            btn_verificar = page.get_by_role("button", name=re.compile(r"^Verificar$", re.I))
            expect(btn_verificar).to_be_visible(timeout=30000)
            # Pode ou não navegar; trate os dois casos
            try:
                click_and_wait_dom(page, btn_verificar)
            except Exception:
                btn_verificar.click()
                page.wait_for_load_state("domcontentloaded")

            # 4) Clicar "Simplificado" (às vezes abre nova aba)
            btn_simplificado = page.get_by_role("button", name=re.compile(r"Simplificado", re.I)).or_(
                page.get_by_text(re.compile(r"Simplificado", re.I))
            )
            expect(btn_simplificado).to_be_visible(timeout=60000)
            try:
                page = ensure_same_page_or_popup(context, btn_simplificado)
            except Exception:
                btn_simplificado.click()
                page.wait_for_load_state("domcontentloaded")

            # 6) Clicar "Folha de Pagamento"
            folha = (
                page.get_by_role("link", name=re.compile(r"Folha de Pagamento", re.I))
                .or_(page.get_by_role("button", name=re.compile(r"Folha de Pagamento", re.I)))
                .or_(page.get_by_text(re.compile(r"Folha de Pagamento", re.I)))
            )
            expect(folha).to_be_visible(timeout=60000)
            try:
                click_and_wait_dom(page, folha)
            except Exception:
                folha.click()
                page.wait_for_load_state("domcontentloaded")

            # 8) Clicar "Encerrar Folha"
            encerrar = (
                page.get_by_role("button", name=re.compile(r"Encerrar Folha", re.I))
                .or_(page.get_by_text(re.compile(r"Encerrar Folha", re.I)))
            )
            expect(encerrar).to_be_visible(timeout=60000)
            try:
                click_and_wait_dom(page, encerrar)
            except Exception:
                encerrar.click()
                page.wait_for_load_state("domcontentloaded")

            # 10) Clicar "Cancelar" (geralmente modal → sem navegação)
            cancelar = page.get_by_role("button", name=re.compile(r"Cancelar", re.I)).or_(
                page.get_by_text(re.compile(r"Cancelar", re.I))
            )
            expect(cancelar).to_be_visible(timeout=30000)
            cancelar.click()
            # Curta espera para estabilidade da UI
            page.wait_for_timeout(800)

            print(f"[OK] CPF processado: {cpf}")

        print("Processo finalizado para todos os CPFs.")