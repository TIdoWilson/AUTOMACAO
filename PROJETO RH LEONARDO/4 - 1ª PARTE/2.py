import argparse
import os
import shutil
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(__file__)
DET_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "3 - DET"))
if DET_DIR not in sys.path:
    sys.path.append(DET_DIR)

from chrome_9222 import PORT, chrome_9222


# =========================
# CONFIG
# =========================

ESOCIAL_URL = "https://www.esocial.gov.br/portal/Home/Index?trocarPerfil=true"

EXCEL_PATH = str(Path(BASE_DIR) / "domesticas_preparado.xlsx")
SHEET_NAME = "Domésticas"
COL_NOME = "Nome"
COL_CPF = "CPF"
COL_OBS = "VT/VR"

DOWNLOAD_DIR = str(Path(BASE_DIR) / "eSocial Site Guias")
FINAL_ROOT = r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\DOMÉSTICAS"
NO_AUTH_PREFIX = "Empregadores sem autorização - "

DEFAULT_WAIT_LOGIN_S = 1.0
VERIFY_DELAY_S = 2.0
MAX_RETRIES_PER_CPF = 3


# =========================
# UTIL
# =========================

def sanitize_filename(name: str) -> str:
    invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid:
        name = name.replace(ch, "-")
    return " ".join(name.split()).strip()


def strip_procuracao(name: str) -> str:
    text = name.strip()
    lowered = normalize_text(text)
    if lowered.endswith(" - procuracao"):
        return text[: text.rfind(" - ")].strip()
    return text


def find_no_auth_log_path() -> Path:
    root = Path(FINAL_ROOT)
    root.mkdir(parents=True, exist_ok=True)
    candidates = list(root.glob(f"{NO_AUTH_PREFIX}*.txt"))
    if not candidates:
        return root / f"{NO_AUTH_PREFIX}{datetime.now():%Y-%m-%d_%H%M%S}.txt"
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def append_no_auth_log(name: str, cpf: str) -> Path:
    log_path = find_no_auth_log_path()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{strip_procuracao(name)} | {cpf} | {timestamp}\n"
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line)
    new_path = Path(FINAL_ROOT) / f"{NO_AUTH_PREFIX}{datetime.now():%Y-%m-%d_%H%M%S}.txt"
    if log_path != new_path:
        try:
            log_path.replace(new_path)
            return new_path
        except Exception:
            return log_path
    return log_path


def normalize_text(value) -> str:
    if value is None:
        raw = ""
    else:
        raw = str(value)
    normalized = unicodedata.normalize("NFKD", raw)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch)).strip().lower()


def ensure_dir(path: str | Path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def resolve_excel_path(excel_path: str) -> str:
    if excel_path and os.path.exists(excel_path):
        return excel_path
    candidates = list(Path(BASE_DIR).glob("*preparado*.xlsx"))
    if not candidates:
        return excel_path
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return str(candidates[0])


def load_companies(excel_path: str, sheet: str) -> list[dict]:
    try:
        import pandas as pd
    except Exception as exc:
        raise SystemExit("pandas is required to read the Excel file") from exc
    excel_path = resolve_excel_path(excel_path)
    if not excel_path or not os.path.exists(excel_path):
        raise SystemExit("Excel path not found. Generate the prepared Excel first.")
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet)
    except ValueError:
        df = pd.read_excel(excel_path, sheet_name=0)
    missing = [c for c in (COL_NOME, COL_CPF) if c not in df.columns]
    if missing:
        raise SystemExit(f"Missing columns in Excel: {', '.join(missing)}")
    items = []
    for _, row in df.iterrows():
        name = str(row.get(COL_NOME, "")).strip()
        cpf = str(row.get(COL_CPF, "")).strip()
        if not name or not cpf:
            continue
        if COL_OBS in df.columns:
            obs = normalize_text(row.get(COL_OBS, ""))
            if "afastada nao tem guia" in obs:
                continue
        items.append({"name": name, "cpf": cpf})
    return items


def click_emitir_guia(page):
    try:
        page.locator("#btn-emitir-guia").click(timeout=10000)
        return
    except Exception:
        pass
    try:
        page.locator("a:has-text('Emitir Guia')").first.click(timeout=10000)
        return
    except Exception:
        pass
    page.locator("button:has-text('Emitir Guia')").first.click(timeout=10000)


def try_emitir_guia(page) -> bool:
    try:
        if page.locator("#btn-emitir-guia").first.is_visible():
            click_emitir_guia(page)
            return True
    except Exception:
        pass
    try:
        if page.locator("a:has-text('Emitir Guia')").first.is_visible():
            click_emitir_guia(page)
            return True
    except Exception:
        pass
    try:
        if page.locator("button:has-text('Emitir Guia')").first.is_visible():
            click_emitir_guia(page)
            return True
    except Exception:
        pass
    return False


def has_no_auth_message(page) -> bool:
    try:
        text = page.locator("#mensagemGeral").inner_text(timeout=1000)
        txt_norm = normalize_text(text)
        if "procurador" in txt_norm and "nao possui perfil" in txt_norm:
            return True
    except Exception:
        pass
    try:
        if page.locator(".alert.alert-danger, .fade-alert.alert.alert-danger").first.is_visible():
            txt = normalize_text(page.locator(".alert.alert-danger, .fade-alert.alert.alert-danger").first.inner_text(timeout=1000))
            if "procurador" in txt and "nao possui perfil" in txt and "autorizacao de acesso a web" in txt:
                return True
    except Exception:
        pass
    try:
        body_text = normalize_text(page.locator("body").inner_text(timeout=1000))
        if "o procurador nao possui perfil com autorizacao de acesso a web" in body_text:
            return True
    except Exception:
        pass
    return False


def safe_click(page, selector: str, timeout_ms: int = 15000, no_wait_after: bool = False) -> None:
    loc = page.locator(selector).first
    try:
        loc.wait_for(state="attached", timeout=timeout_ms)
    except Exception:
        pass
    try:
        loc.scroll_into_view_if_needed()
    except Exception:
        pass
    try:
        loc.click(timeout=timeout_ms, no_wait_after=no_wait_after)
        return
    except Exception:
        pass
    try:
        page.evaluate(
            "(sel) => { const el = document.querySelector(sel); if (el) el.click(); }",
            selector,
        )
    except Exception:
        loc.click(force=True, timeout=timeout_ms, no_wait_after=no_wait_after)


def wait_for_download_and_save(page, download_dir: Path, target_name: str) -> Path:
    download_dir.mkdir(parents=True, exist_ok=True)
    with page.expect_download(timeout=60000) as dl_info:
        click_emitir_guia(page)
    download = dl_info.value
    tmp_name = sanitize_filename(target_name)
    tmp_path = download_dir / tmp_name
    if tmp_path.suffix.lower() != ".pdf":
        tmp_path = tmp_path.with_suffix(".pdf")
    download.save_as(str(tmp_path))
    return tmp_path


def move_to_final(pdf_path: Path, name: str) -> Path:
    now = datetime.now()
    year = str(now.year)
    month = f"{now.month:02d}"
    company = sanitize_filename(strip_procuracao(name))
    final_dir = ensure_dir(Path(FINAL_ROOT) / year / month / company)
    final_name = f"DAE - {company}.pdf"
    dest = final_dir / final_name
    if dest.exists():
        idx = 2
        while True:
            alt = final_dir / f"DAE - {company} ({idx}).pdf"
            if not alt.exists():
                dest = alt
                break
            idx += 1
    shutil.copy2(pdf_path, dest)
    return dest


def has_existing_guide(name: str) -> bool:
    now = datetime.now()
    year = str(now.year)
    month = f"{now.month:02d}"
    company = sanitize_filename(strip_procuracao(name))
    final_dir = Path(FINAL_ROOT) / year / month / company
    if not final_dir.exists():
        return False
    prefix = f"DAE - {company}"
    return any(p.is_file() and p.name.startswith(prefix) for p in final_dir.glob("*.pdf"))


def run_flow(wait_login_s: float, headless: bool = False):
    with sync_playwright() as p:
        if headless:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(accept_downloads=True)
        else:
            browser = chrome_9222(p, PORT)
            context = browser.contexts[0] if browser.contexts else browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.goto(ESOCIAL_URL)
        time.sleep(wait_login_s)
        try:
            page.locator("button:has-text('Entrar com'):has-text('gov.br')").first.click(timeout=10000)
        except Exception:
            try:
                page.locator("button:has-text('Entrar com')").first.click(timeout=10000)
            except Exception:
                pass
        page.wait_for_load_state("domcontentloaded")
        try:
            page.locator("#login-certificate").first.click(timeout=10000)
        except Exception:
            pass
        if headless:
            print("[info] modo headless: aguardando login automatico.")
            time.sleep(max(wait_login_s, 5.0))
        else:
            print("[info] aguarde o login no site.")
            input("[acao] Depois de logar e confirmar o acesso, pressione ENTER para continuar.")

        companies = load_companies(EXCEL_PATH, SHEET_NAME)
        if not companies:
            print("[info] nenhuma empresa encontrada no excel.")
            return

        download_dir = ensure_dir(DOWNLOAD_DIR)

        for item in companies:
            name = item["name"]
            cpf = item["cpf"]

            if has_existing_guide(name):
                print(f"[skip] guia ja existe: {name}")
                continue

            attempt = 0
            while attempt < MAX_RETRIES_PER_CPF:
                attempt += 1
                try:
                    try:
                        page.goto(ESOCIAL_URL, wait_until="domcontentloaded")
                    except Exception:
                        page.goto("https://www.esocial.gov.br/portal/", wait_until="domcontentloaded")
                    page.select_option("#perfilAcesso", "PROCURADOR_PF")
                    page.fill("#procuradorCpf", cpf)
                    page.click("#btn-verificar-procuracao-cpf")
                    page.wait_for_timeout(int(VERIFY_DELAY_S * 1000))
                    if has_no_auth_message(page):
                        log_path = append_no_auth_log(name, cpf)
                        print(f"[erro] sem autorizacao: {name} -> {log_path}")
                        break

                    safe_click(page, "#domestico:has-text('Simplificado')", timeout_ms=20000)
                    page.wait_for_load_state("domcontentloaded")
                    safe_click(page, "a.card-acesso.folha-pagamento:has-text('Folha de Pagamento')", timeout_ms=20000)
                    page.wait_for_load_state("domcontentloaded")
                    try:
                        page.locator("#btnContinuar").click(timeout=8000)
                    except Exception:
                        pass

                    display_name = strip_procuracao(name)
                    if try_emitir_guia(page):
                        tmp_pdf = wait_for_download_and_save(page, download_dir, f"DAE - {display_name}.pdf")
                    else:
                        safe_click(page, "#btnEncerrarFolha", timeout_ms=12000, no_wait_after=True)
                        safe_click(page, "#btnConfirmar", timeout_ms=12000, no_wait_after=True)
                        tmp_pdf = wait_for_download_and_save(page, download_dir, f"DAE - {display_name}.pdf")
                    final_path = move_to_final(tmp_pdf, display_name)
                    print(f"[ok] salvo: {final_path}")
                    break
                except Exception as exc:
                    try:
                        if has_no_auth_message(page):
                            log_path = append_no_auth_log(name, cpf)
                            print(f"[erro] sem autorizacao (detectado em excecao): {name} -> {log_path}")
                            break
                    except Exception:
                        pass
                    print(f"[warn] falha ({attempt}/{MAX_RETRIES_PER_CPF}) para {name}: {exc}")
                    if attempt >= MAX_RETRIES_PER_CPF:
                        print(f"[erro] desistindo: {name}")
                    else:
                        time.sleep(1.5)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--wait-login", type=float, default=DEFAULT_WAIT_LOGIN_S)
    ap.add_argument("--headless", action="store_true")
    args = ap.parse_args()
    run_flow(wait_login_s=args.wait_login, headless=args.headless)


if __name__ == "__main__":
    main()

