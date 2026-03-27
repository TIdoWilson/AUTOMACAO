import time
from datetime import date, datetime
from calendar import monthrange
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
from chrome_9222 import PORT, chrome_9222


LOGIN_URL = "https://cav.receita.fazenda.gov.br/autenticacao/login"
APLICACAO_URL = "https://cav.receita.fazenda.gov.br/ecac/Aplicacao.aspx?id=10015&origem=maisacessados"

DOWNLOAD_DIR = r"T:\testes"
OUTPUT_EXCEL_NAME = "recibos_ecac.xlsx"

TIMEOUT_MS = 30_000
MAX_DOWNLOAD_RETRIES = 5


def first_day_last_month(today: date) -> date:
    if today.month == 1:
        return date(today.year - 1, 12, 1)
    return date(today.year, today.month - 1, 1)


def last_day_current_month(today: date) -> date:
    last_day = monthrange(today.year, today.month)[1]
    return date(today.year, today.month, last_day)


def fmt_br(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def find_in_frames(page, selector: str, timeout_ms: int = TIMEOUT_MS):
    end = time.time() + timeout_ms / 1000
    last_err = None
    while time.time() < end:
        for frame in page.frames:
            loc = frame.locator(selector)
            try:
                if loc.count() > 0:
                    return frame, loc
            except Exception as exc:
                last_err = exc
        time.sleep(0.2)
    raise PWTimeoutError(f"Selector nao encontrado: {selector}. Erro: {last_err}")


def wait_processing_end(frame, timeout_ms: int = TIMEOUT_MS):
    end = time.time() + timeout_ms / 1000
    while time.time() < end:
        try:
            proc = frame.locator(
                "img.image-processamento, img[class*='image-processamento']"
            )
            if proc.count() == 0:
                return
            if not proc.first.is_visible():
                return
        except Exception:
            return
        time.sleep(0.2)


def click_blank_area(frame):
    try:
        frame.locator("body").click(position={"x": 5, "y": 5})
    except Exception:
        pass


def wait_for_login(page):
    page.goto(LOGIN_URL)
    try:
        page.locator("input[alt='Acesso Gov BR'], input[src*='gov-br.png']").first.click(
            timeout=8000
        )
        page.wait_for_load_state("networkidle")
        time.sleep(1)
        try:
            page.locator("#login-certificate").first.click(timeout=3000)
        except Exception:
            pass
    except Exception:
        pass
    print("[info] aguarde o login no site.")
    input("[acao] Depois de logar e confirmar o acesso, pressione ENTER para continuar.")
    try:
        page.wait_for_url("**cav.receita.fazenda.gov.br/**", timeout=20000)
    except Exception:
        pass


def current_page_number(page) -> str | None:
    try:
        _, loc = find_in_frames(
            page, "a.ui-paginator.aspNetDisabled, a.aspNetDisabled.ui-paginator"
        )
        txt = loc.first.inner_text().strip()
        return txt if txt else None
    except Exception:
        return None


def read_table_headers(frame):
    header_cells = frame.locator(
        "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs thead th"
    )
    headers = [h.strip() for h in header_cells.all_inner_texts()]
    if headers:
        return headers
    header_cells = frame.locator(
        "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs tr th"
    )
    headers = [h.strip() for h in header_cells.all_inner_texts()]
    return headers


def read_table_rows(frame, headers):
    rows = []
    tr_loc = frame.locator(
        "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs tbody tr"
    )
    if tr_loc.count() == 0:
        tr_loc = frame.locator(
            "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs tr"
        )
    for i in range(tr_loc.count()):
        tr = tr_loc.nth(i)
        if tr.locator("th").count() > 0:
            continue
        cells = [c.strip() for c in tr.locator("td").all_inner_texts()]
        if not cells:
            continue
        if not headers:
            headers = [f"col_{idx+1}" for idx in range(len(cells))]
        if len(cells) < len(headers):
            cells = cells + [""] * (len(headers) - len(cells))
        rows.append((tr, dict(zip(headers, cells))))
    return rows


def extract_doc_from_row(row_dict):
    for key in row_dict:
        k = key.lower()
        if "cnpj" in k or "cpf" in k or "identifica" in k:
            return row_dict.get(key, "").strip()
    return None


def main():
    base_dir = Path(__file__).resolve().parent
    download_dir = Path(DOWNLOAD_DIR)
    download_dir.mkdir(parents=True, exist_ok=True)
    output_path = base_dir / OUTPUT_EXCEL_NAME

    today = date.today()
    data_ini = fmt_br(first_day_last_month(today))
    data_fim = fmt_br(last_day_current_month(today))

    with sync_playwright() as p:
        browser = chrome_9222(p, PORT)
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = context.new_page()

        wait_for_login(page)

        frame, chk = find_in_frames(page, "#ctl00_cphConteudo_chkListarOutorgantes")
        chk.click()
        time.sleep(60)

        frame, dropdown_btn = find_in_frames(
            page, "button[data-id='ctl00_cphConteudo_ddlOutorgantes']"
        )
        dropdown_btn.click()
        frame.locator(".dropdown-menu.open .bs-searchbox input.form-control").fill("CNPJ")
        frame.locator(".dropdown-menu.open button.actions-btn.bs-select-all").first.click()
        dropdown_btn.click()

        frame, di = find_in_frames(page, "#txtDataInicio")
        frame.locator("#txtDataInicio").fill(data_ini)
        frame.locator("#txtDataFinal").fill(data_fim)
        frame.locator("#txtDataTransmissaoInicial").fill(data_ini)
        try:
            frame.keyboard.press("Escape")
        except Exception:
            pass
        click_blank_area(frame)

        frame.locator("#ctl00_cphConteudo_imageFiltrar").click()
        wait_processing_end(frame)

        frame, tabela = find_in_frames(
            page, "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"
        )
        tabela.wait_for(state="visible", timeout=TIMEOUT_MS)
        frame.locator(
            "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs tr"
        ).first.wait_for(state="attached", timeout=TIMEOUT_MS)

        headers = read_table_headers(frame)
        all_rows = []
        downloaded_docs = []

        while True:
            page_num = current_page_number(page)
            rows = read_table_rows(frame, headers)
            if not headers and rows:
                headers = list(rows[0][1].keys())

            expected_downloads = 0
            for tr, _ in rows:
                if tr.locator("a[title='Visualizar Recibo']").count() > 0:
                    expected_downloads += 1
            print(f"[info] pagina {page_num or '?'}: {expected_downloads} recibos.")

            downloaded_this_page = 0
            for tr, row_dict in rows:
                all_rows.append(row_dict)
                doc = extract_doc_from_row(row_dict)
                link = tr.locator("a[title='Visualizar Recibo']")
                if link.count() > 0:
                    attempts = 0
                    while True:
                        try:
                            with page.expect_download(timeout=TIMEOUT_MS) as dl_info:
                                link.first.click()
                            download = dl_info.value
                            filename = download.suggested_filename
                            download.save_as(str(download_dir / filename))
                            downloaded_docs.append(
                                {
                                    "documento": doc or "",
                                    "arquivo": filename,
                                    "pagina": page_num or "",
                                }
                            )
                            downloaded_this_page += 1
                            click_blank_area(frame)
                            break
                        except PWTimeoutError:
                            attempts += 1
                            wait_processing_end(frame)
                            if attempts >= MAX_DOWNLOAD_RETRIES:
                                raise PWTimeoutError(
                                    f"Download nao iniciou apos {MAX_DOWNLOAD_RETRIES} tentativas."
                                )

            if downloaded_this_page < expected_downloads:
                raise RuntimeError(
                    f"Baixados {downloaded_this_page} de {expected_downloads} na pagina {page_num or '?'}."
                )

            _, next_link = find_in_frames(
                page, "#ctl00_cphConteudo_tabelaListagemDctf_paginacaoListagemDctf_lnkNextPage"
            )
            class_attr = next_link.first.get_attribute("class") or ""
            if "aspNetDisabled" in class_attr:
                break

            first_row_text = ""
            if rows:
                try:
                    first_row_text = rows[0][0].inner_text()
                except Exception:
                    first_row_text = ""

            next_link.first.click()
            wait_processing_end(frame)
            frame, tabela = find_in_frames(
                page, "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"
            )

            if first_row_text:
                try:
                    frame.locator(
                        "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs tbody tr"
                    ).first.wait_for(state="attached", timeout=TIMEOUT_MS)
                    frame.wait_for_function(
                        """(sel, oldText) => {
                            const el = document.querySelector(sel);
                            if (!el) return false;
                            return el.innerText !== oldText;
                        }""",
                        (
                            "#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs tbody tr",
                            first_row_text,
                        ),
                        timeout=TIMEOUT_MS,
                    )
                except Exception:
                    pass

        df = pd.DataFrame(all_rows, columns=headers)
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, index=False, sheet_name="tabela")
            if downloaded_docs:
                pd.DataFrame(downloaded_docs).to_excel(writer, index=False, sheet_name="baixados")

        print(f"Arquivos salvos em: {download_dir}")
        print(f"Planilha salva em: {output_path}")

        context.close()
        browser.close()


if __name__ == "__main__":
    main()

