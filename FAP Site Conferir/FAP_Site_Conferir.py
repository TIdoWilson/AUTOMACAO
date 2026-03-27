import csv
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).resolve().parent
DET_DIR = BASE_DIR.parent / "PROJETO RH LEONARDO" / "3 - DET"
if DET_DIR.exists() and str(DET_DIR) not in sys.path:
    sys.path.append(str(DET_DIR))

from chrome_9222 import PORT, chrome_9222
DOWNLOAD_DIR = BASE_DIR / "downloads"
LOG_PATH = BASE_DIR / "download_log.csv"

URL = "https://fap-mps.dataprev.gov.br/"
VIGENCIA = "2026"

WAIT_LONG = 120_000
WAIT_MED = 30_000
WAIT_SHORT = 5_000

DOWNLOAD_SELECTORS = [
    "button[title='Download PDF']",
    "text=Baixar PDF",
    "text=Download PDF",
    "text=Baixar",
    "button:has-text('PDF')",
    "a:has-text('PDF')",
]


def _normalize_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def _sanitize_filename(name: str) -> str:
    invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid:
        name = name.replace(ch, "-")
    return " ".join(name.split()).strip()


def _append_log(cnpj_raiz: str, company_name: str, status: str = "ok", detail: str = "") -> None:
    is_new = not LOG_PATH.exists()
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        if is_new:
            writer.writerow(["timestamp", "cnpj_raiz", "nome", "status", "detail"])
        writer.writerow(
            [datetime.now().strftime("%Y-%m-%d %H:%M:%S"), cnpj_raiz, company_name, status, detail]
        )


def _open_dropdown(input_loc) -> None:
    input_loc.click()
    btn = input_loc.locator("xpath=following-sibling::button[1]")
    if btn.count() > 0:
        btn.first.click()
    try:
        input_loc.page.wait_for_timeout(300)
    except Exception:
        pass


def _list_open_options(page) -> list[str]:
    selectors = [
        "ul.MuiAutocomplete-listbox li",
        "div.MuiAutocomplete-popper li",
        "[role='listbox'] [role='option']",
        "ul[role='listbox'] li",
        ".MuiAutocomplete-listbox li",
        ".br-list [role='option']",
        ".br-list li",
        ".MuiPaper-root li",
    ]
    for sel in selectors:
        loc = page.locator(sel)
        if loc.count() == 0:
            continue
        texts = [t.strip() for t in loc.all_inner_texts() if t.strip()]
        if texts:
            seen = set()
            unique = []
            for t in texts:
                if t not in seen:
                    unique.append(t)
                    seen.add(t)
            return unique
    return []


def _select_autocomplete(page, input_id: str, value: str) -> None:
    input_loc = page.locator(f"#{input_id}")
    input_loc.click()
    input_loc.fill(value)
    try:
        opt = page.locator("[role='option']", has_text=value)
        if opt.count() > 0:
            opt.first.click()
            return
    except Exception:
        pass
    page.keyboard.press("Enter")


def _company_name_from_cnpj_option(text: str) -> str:
    raw = (text or "").strip()
    if not raw:
        return ""
    parts = [p.strip() for p in raw.split("-", 1)]
    if len(parts) == 2 and parts[1]:
        return parts[1]
    stripped = _normalize_digits(raw)
    if stripped and stripped in raw:
        name = raw.replace(stripped, "").strip(" -/\\")
        if name:
            return name
    return ""


def _get_cnpj_options(page) -> list[str]:
    input_loc = page.locator("#cnpjRaiz")
    _open_dropdown(input_loc)
    time.sleep(0.5)
    options = _list_open_options(page)
    if options:
        return options
    input_loc.click()
    page.keyboard.press("ArrowDown")
    time.sleep(0.5)
    return _list_open_options(page)


def _wait_for_consulta_card(page) -> None:
    page.locator("a[href='/consultar-fap']").first.wait_for(state="visible", timeout=WAIT_LONG)


def _click_if_visible(page, selector: str) -> bool:
    loc = page.locator(selector)
    if loc.count() == 0:
        return False
    try:
        loc.first.click()
        return True
    except Exception:
        return False


def _extract_company_name(page) -> str:
    labels = ["Razão Social", "Razao Social", "Nome Empresarial", "Nome"]
    for label in labels:
        label_loc = page.locator(f"text={label}").first
        if label_loc.count() == 0:
            continue
        try:
            value = label_loc.locator("xpath=following::*[1]").inner_text().strip()
            if value:
                return value
        except Exception:
            continue
    return ""


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for i in range(2, 1000):
        candidate = path.with_name(f"{stem} ({i}){suffix}")
        if not candidate.exists():
            return candidate
    return path


def _existing_downloads_count(company_name: str) -> int:
    if not DOWNLOAD_DIR.exists():
        return 0
    pattern = f"{company_name} - *.pdf"
    return len(list(DOWNLOAD_DIR.glob(pattern)))


def _download_exists_for_index(company_name: str, idx: int) -> bool:
    return (DOWNLOAD_DIR / f"{company_name} - {idx}.pdf").exists()


def _wait_for_download_button(page) -> str | None:
    for sel in DOWNLOAD_SELECTORS:
        try:
            page.locator(sel).first.wait_for(state="visible", timeout=WAIT_MED)
            return sel
        except PWTimeout:
            continue
    return None


def _has_estabelecimento_inexistente(page) -> bool:
    msg = "Estabelecimento inexistente ou não constante na base de dados do FAP vigência 2026"
    loc = page.locator(".br-message.danger", has_text=msg)
    try:
        return loc.first.is_visible()
    except Exception:
        return False


def _download_pdf(page, file_path: Path) -> None:
    selector = _wait_for_download_button(page)
    if not selector:
        raise RuntimeError("Botao de download nao apareceu a tempo.")
    for sel in DOWNLOAD_SELECTORS:
        try:
            with page.expect_download(timeout=WAIT_MED) as dl_info:
                clicked = _click_if_visible(page, sel)
                if not clicked:
                    continue
            download = dl_info.value
            download.save_as(str(file_path))
            return
        except PWTimeout:
            continue
    raise RuntimeError("Nao foi possivel iniciar o download do PDF.")


def _ensure_download_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def run() -> None:
    _ensure_download_dir(DOWNLOAD_DIR)

    with sync_playwright() as p:
        browser = chrome_9222(p, PORT)
        context = browser.contexts[0] if browser.contexts else browser.new_context(accept_downloads=True)
        page = context.new_page()

        print("Abrindo site FAP...")
        page.goto(URL, wait_until="domcontentloaded")

        print("Clicando em 'Entrar com gov.br'...")
        try:
            page.locator("button:has-text('Entrar com'):has-text('gov.br')").first.click(timeout=10000)
        except Exception:
            page.locator("button.br-button.primary", has_text="Entrar com gov.br").first.click()

        print("Selecionando 'Seu certificado digital'...")
        try:
            page.locator("#login-certificate").first.click(timeout=10000)
        except Exception:
            pass

        print("[info] aguarde o login com certificado concluir.")
        input("[acao] Depois de logar e confirmar o acesso, pressione ENTER para continuar.")
        _wait_for_consulta_card(page)

        print("Abrindo Consulta do FAP...")
        try:
            page.locator("a[title*='Consulte o valor do FAP']").first.click(timeout=10000)
        except Exception:
            page.locator("a[href='/consultar-fap']").last.click(timeout=10000)

        print(f"Selecionando vigencia {VIGENCIA}...")
        _select_autocomplete(page, "anoVigencia", VIGENCIA)

        cnpj_options = _get_cnpj_options(page)
        if not cnpj_options:
            raise RuntimeError("Nenhum CNPJ raiz encontrado no site.")

        for cnpj_text in cnpj_options:
            cnpj_raiz = _normalize_digits(cnpj_text)[:8] or cnpj_text
            company_base = _company_name_from_cnpj_option(cnpj_text)
            company_name = _sanitize_filename(company_base or "Empresa")
            try:
                print(f"Processando CNPJ raiz {cnpj_text}...")
                _select_autocomplete(page, "cnpjRaiz", cnpj_text)
                time.sleep(3)

                est_input = page.locator("#estabelecimentos")
                _open_dropdown(est_input)
                time.sleep(0.5)
                estabelecimentos = _list_open_options(page)
                if not estabelecimentos:
                    est_input.click()
                    page.keyboard.press("ArrowDown")
                    time.sleep(0.5)
                    estabelecimentos = _list_open_options(page)
                if not estabelecimentos:
                    if _has_estabelecimento_inexistente(page):
                        company_name = _sanitize_filename(
                            company_base or _extract_company_name(page) or "Empresa"
                        )
                        _append_log(
                            cnpj_raiz,
                            company_name,
                            status="erro",
                            detail="estabelecimento inexistente na vigencia 2026",
                        )
                        print("[erro] estabelecimento inexistente. Pulando CNPJ.")
                        continue
                    raise RuntimeError(f"Nenhum estabelecimento encontrado para {cnpj_text}.")

                company_name = _sanitize_filename(company_base or _extract_company_name(page) or "Empresa")
                total_existing = _existing_downloads_count(company_name)
                if total_existing >= len(estabelecimentos):
                    print(f"[skip] ja existem {total_existing} PDFs para {company_name}.")
                    continue

                for idx, est in enumerate(estabelecimentos, start=1):
                    print(f"- Estabelecimento {idx}/{len(estabelecimentos)}: {est}")
                    if _download_exists_for_index(company_name, idx):
                        print(f"[skip] PDF ja existe para {company_name} - {idx}.")
                        continue
                    _select_autocomplete(page, "estabelecimentos", est)

                    print("Consultando...")
                    page.get_by_role("button", name="Consultar").click()
                    time.sleep(3)

                    attempts = 0
                    while _has_estabelecimento_inexistente(page) and attempts < 3:
                        attempts += 1
                        print(f"[warn] estabelecimento inexistente. Tentativa {attempts}/3.")
                        _select_autocomplete(page, "cnpjRaiz", cnpj_text)
                        time.sleep(2)
                        _select_autocomplete(page, "estabelecimentos", est)
                        page.get_by_role("button", name="Consultar").click()
                        time.sleep(3)
                    if _has_estabelecimento_inexistente(page):
                        _append_log(
                            cnpj_raiz,
                            company_name,
                            status="erro",
                            detail="estabelecimento inexistente na vigencia 2026",
                        )
                        print("[erro] estabelecimento inexistente persistente. Pulando CNPJ.")
                        break

                    file_name = f"{company_name} - {idx}.pdf"
                    target = _unique_path(DOWNLOAD_DIR / file_name)

                    print(f"Baixando PDF: {target.name}")
                    _download_pdf(page, target)
                    _append_log(cnpj_raiz, company_name, status="ok")
            except Exception as exc:
                _append_log(cnpj_raiz, company_name, status="erro", detail=str(exc))
                print(f"[erro] falha no CNPJ {cnpj_text}: {exc}")
                continue

        print("Concluido.")
        context.close()
        browser.close()


if __name__ == "__main__":
    run()
