import os
import re
import glob
import time
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# =====================================
# Configurações de URL (Empresas)
# =====================================
URL = "https://web.tareffa.com.br/empresas"

# =====================================
# Helpers de espera
# =====================================

LOGIN_EMAIL = "robos@wilsonlopes.com.br"
LOGIN_password = "123456"

def wait_dom(page):
    page.wait_for_load_state("domcontentloaded")


def wait_new_empresa_form_ready(page):
    """
    Garante que o formulário de 'Nova Empresa' foi renderizado.
    Considera campos típicos: Código ERP, Nome Fantasia, Razão Social.
    """
    seletor = (
        'input[placeholder*="Código ERP" i], '
        'input[placeholder*="Nome Fantasia" i], '
        'input[placeholder*="Razão Social" i], '
        'mat-form-field:has(mat-label:has-text("Código ERP")) input, '
        'mat-form-field:has(mat-label:has-text("Nome Fantasia")) input, '
        'mat-form-field:has(mat-label:has-text("Razão Social")) input'
    )
    page.wait_for_selector(seletor, timeout=45000)


# =====================================
# Preenchimento de campos
# =====================================
# =========================
# Login
# =========================
def fill_email(page, value):
    candidatos = [
        page.get_by_label(re.compile(r"e-?mail", re.I)),
        page.get_by_placeholder(re.compile(r"e-?mail", re.I)),
        page.locator('input[type="email"]'),
        page.locator('input[name="email"]'),
        page.locator('input[formcontrolname="email"]'),
        page.locator('input[autocomplete="username"]'),
    ]
    for loc in candidatos:
        if loc.count():
            try:
                loc.first.wait_for(state="visible")
                loc.first.fill(value)
                return
            except Exception:
                continue
    raise RuntimeError("Campo de e-mail não encontrado.")

def fill_password(page, value):
    candidatos = [
        page.get_by_label(re.compile(r"senha|password", re.I)),
        page.get_by_placeholder(re.compile(r"senha|password", re.I)),
        page.locator('input[type="password"]'),
        page.locator('input[name="password"]'),
        page.locator('input[formcontrolname="password"]'),
        page.locator('input[autocomplete="current-password"]'),
    ]
    for loc in candidatos:
        if loc.count():
            try:
                loc.first.wait_for(state="visible")
                loc.first.fill(value)
                return
            except Exception:
                continue
    raise RuntimeError("Campo de senha não encontrado.")

def click_entrar(page):
    candidatos = [
        page.get_by_role("button", name=re.compile(r"^\s*entrar\s*$", re.I)),
        page.locator("button").filter(has_text=re.compile(r"\bentrar\b", re.I)),
        page.locator('button[type="submit"]'),
    ]
    for loc in candidatos:
        if loc.count():
            try:
                loc.first.wait_for(state="visible")
                loc.first.click()
                return
            except Exception:
                continue
    raise RuntimeError("Botão 'ENTRAR' não encontrado.")

def click_overlay_submit(page):
    # Botão absoluto com z-index (se existir)
    alvo = page.locator(
        'button[type="submit"][style*="position: absolute"][style*="z-index: 10"]'
    )
    if alvo.count():
        try:
            alvo.first.wait_for(state="visible")
            alvo.first.click()
            return
        except Exception:
            pass
    # Fallback: qualquer submit
    comum = page.locator('button[type="submit"]')
    if comum.count():
        try:
            comum.first.wait_for(state="visible")
            comum.first.click()
            return
        except Exception:
            pass
    raise RuntimeError("Botão submit (overlay) não encontrado.")

def _count_safe(locator):
    try:
        return locator.count()
    except Exception:
        return 0

def _fill_first(page, locator, value):
    if _count_safe(locator):
        el = locator.first
        try:
            el.scroll_into_view_if_needed()
        except Exception:
            pass
        el.wait_for(state="visible", timeout=8000)
        el.fill(value)
        return True
    return False