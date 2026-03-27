# -*- coding: utf-8 -*-
"""
DCTFWEB 1.2.6 (teste_fast) — Playwright + CDP attach (Chrome 9222)
Objetivo: baixar recibos MAIS RÁPIDO mantendo a lógica "múltiplos ao mesmo tempo".

Estratégia principal (rápida):
- Clica em lote com no_wait_after=True
- Captura respostas HTTP PDF no BrowserContext (context.on("response"))
- Salva PDF direto em disco (sem expect_download e sem polling de .crdownload)

Fallback (para os que não vierem como resposta PDF):
- Ativa download via CDP (Page.setDownloadBehavior)
- Clica e aguarda arquivo aparecer na pasta (polling), em pequenos lotes

Pré-req:
  pip install playwright
  playwright install chromium

Execução:
  - Abra o Chrome com --remote-debugging-port=9222 (já logado)
  - Deixe a tela da DCTFWeb aberta com a tabela visível
  - Rode este script
"""

import os
import re
import time
import hashlib
from pathlib import Path
from urllib.parse import unquote
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

# ===================== CONFIG =====================
CHROME_DEBUG_ADDR = os.environ.get("CHROME_DEBUG_ADDR", "http://localhost:9222")
DOWNLOAD_DIR = os.environ.get("DCTF_DOWNLOAD_DIR", r"T:\testes")

# Seletor da tabela
SEL_TABELA = "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"

# Link do recibo dentro de cada linha
SEL_LINK_RECIBO = "a[title='Visualizar Recibo'], a[id*='lkbVisualizarRecibo'], a.image-tabela-visualizar-recibo"

# Botão Filtrar (opcional)
SEL_FILTRAR = "a#ctl00_cphConteudo_imageFiltrar, a[id*='imageFiltrar'], a.btn.btn-icon-sec.btn-exibe-processando.image-pesquisar"

# Performance
BATCH_SIZE = int(os.environ.get("DCTF_BATCH_SIZE", "6"))          # tamanho do lote
CLICK_DELAY_MS = int(os.environ.get("DCTF_CLICK_DELAY_MS", "40")) # delay entre cliques no lote
BATCH_TIMEOUT_S = int(os.environ.get("DCTF_BATCH_TIMEOUT_S", "35"))# timeout por lote (captura PDF)
FALLBACK_TIMEOUT_S = int(os.environ.get("DCTF_FALLBACK_TIMEOUT_S", "60"))

# ==================================================

def filename_from_headers(headers: dict, fallback: str) -> str:
    cd = (headers.get("content-disposition") or headers.get("Content-Disposition") or "")
    m = re.search(r"filename\*?=(?:UTF-8''|\"?)([^\";]+)", cd, flags=re.I)
    if m:
        name = unquote(m.group(1)).strip().strip('"')
        return name
    return fallback

def is_pdf_response(resp) -> bool:
    try:
        ct = (resp.headers.get("content-type") or "").lower()
        cd = (resp.headers.get("content-disposition") or "").lower()
        url = (resp.url or "").lower()
        if "application/pdf" in ct:
            return True
        if "filename=" in cd and ".pdf" in cd:
            return True
        if url.endswith(".pdf"):
            return True
        return False
    except Exception:
        return False

def set_download_dir_cdp(page, download_dir: str):
    Path(download_dir).mkdir(parents=True, exist_ok=True)
    client = page.context.new_cdp_session(page)
    try:
        client.send("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": download_dir})
        return
    except Exception:
        pass
    # fallback
    client.send("Browser.setDownloadBehavior", {"behavior": "allow", "downloadPath": download_dir})

def stable_files(path_dir: str):
    try:
        allf = os.listdir(path_dir)
    except FileNotFoundError:
        return set()
    return {f for f in allf if not f.endswith(".crdownload") and not f.endswith(".tmp") and not f.endswith(".partial")}

def has_partials(path_dir: str):
    try:
        return any(n.endswith(".crdownload") or n.endswith(".tmp") or n.endswith(".partial") for n in os.listdir(path_dir))
    except FileNotFoundError:
        return False

def wait_new_files(path_dir: str, before_set: set, expected_min: int, timeout_s: int = 60, quiet_checks: int = 3):
    end = time.time() + timeout_s
    quiet_ok = 0
    while time.time() < end:
        now = stable_files(path_dir)
        new = now - before_set
        if len(new) >= expected_min and not has_partials(path_dir):
            quiet_ok += 1
            if quiet_ok >= quiet_checks:
                return list(new)
        else:
            quiet_ok = 0
        time.sleep(0.25)
    raise PWTimeoutError(f"Timeout esperando {expected_min} novos arquivos em '{path_dir}'")

def wait_visible_any_frame(page, selector: str, timeout_ms=60000):
    end = time.time() + (timeout_ms / 1000)
    last = None
    while time.time() < end:
        for fr in page.frames:
            try:
                # query_selector não espera e é bem rápido
                el = fr.query_selector(selector)
                if el:
                    loc = fr.locator(selector).first
                    loc.wait_for(state="visible", timeout=1000)
                    return fr, loc
            except Exception as e:
                last = e
        time.sleep(0.1)
    raise PWTimeoutError(f"Timeout aguardando: {selector}. Último erro: {last}")

def coletar_links_recibo_da_pagina(page):
    fr, tbl = wait_visible_any_frame(page, SEL_TABELA, timeout_ms=60000)
    rows = tbl.locator("tr")
    n = rows.count()
    links = []
    for i in range(1, n):  # pula cabeçalho
        tr = rows.nth(i)
        link = tr.locator(SEL_LINK_RECIBO).first
        try:
            if link.count() > 0:
                links.append(link)
        except Exception:
            pass
    return links

def selecionar_aba_com_tabela(browser):
    """
    Seleciona a aba que realmente contém a tabela (mesmo que esteja dentro de iframe).
    """
    candidates = []
    for ctx in browser.contexts:
        for pg in ctx.pages:
            candidates.append(pg)

    if not candidates:
        raise RuntimeError("Conectou no CDP mas não há abas visíveis. Verifique a instância/porta 9222.")

    # debug rápido
    print("Abas vistas via CDP:")
    for i, pg in enumerate(candidates):
        try:
            title = (pg.title() or "").strip()
        except Exception:
            title = ""
        try:
            url = pg.url or ""
        except Exception:
            url = ""
        print(f"[{i:02d}] title='{title}' url='{url}'")

    # 1) Melhor: página onde o selector da tabela existe em algum frame
    for pg in candidates:
        try:
            for fr in pg.frames:
                try:
                    if fr.query_selector(SEL_TABELA):
                        return pg
                except Exception:
                    pass
        except Exception:
            pass

    # 2) Fallback: algo que pareça dctfweb/ecac no url/title
    for pg in candidates:
        hay = ""
        try: hay += (pg.url or "")
        except Exception: pass
        try: hay += " " + (pg.title() or "")
        except Exception: pass
        if "dctf" in hay.lower() or "ecac" in hay.lower():
            return pg

    # 3) Último fallback: última aba
    return candidates[-1]

class PdfCollector:
    """
    Coleta PDFs via eventos de response no BrowserContext.
    Mantém dedupe por hash para evitar duplicatas.
    """
    def __init__(self, download_dir: str):
        self.download_dir = download_dir
        Path(download_dir).mkdir(parents=True, exist_ok=True)
        self.saved = []          # paths
        self.errors = []
        self._seen_hashes = set()
        self._counter = 0

    def on_response(self, resp):
        try:
            if not is_pdf_response(resp):
                return
            data = resp.body()
            if not data or not data.startswith(b"%PDF"):
                return

            h = hashlib.sha1(data).hexdigest()
            if h in self._seen_hashes:
                return
            self._seen_hashes.add(h)

            self._counter += 1
            name = filename_from_headers(resp.headers, f"recibo_{self._counter:05d}.pdf")
            if not name.lower().endswith(".pdf"):
                name += ".pdf"

            # evita sobrescrita
            out = Path(self.download_dir) / name
            if out.exists():
                out = Path(self.download_dir) / f"{out.stem}_{self._counter:05d}{out.suffix}"

            out.write_bytes(data)
            self.saved.append(str(out))
        except Exception as e:
            self.errors.append(str(e))

def baixar_recibos_multi_fast(page, download_dir: str, batch_size: int = 6):
    """
    Método rápido:
    - registra listener no context (1 vez)
    - clica em lote com no_wait_after=True
    - aguarda salvar PDFs via responses
    """
    collector = PdfCollector(download_dir)
    ctx = page.context
    ctx.on("response", collector.on_response)

    links = coletar_links_recibo_da_pagina(page)
    if not links:
        print("Nenhum link de recibo encontrado na página.")
        return 0, []

    total = len(links)
    print(f"Links de recibo encontrados: {total}")

    # cliques em lotes
    start_saved = len(collector.saved)
    clicked_total = 0

    for base in range(0, total, batch_size):
        lote = links[base:base+batch_size]
        prev = len(collector.saved)

        # dispara cliques rápido (sem auto-wait)
        clicked = 0
        for link in lote:
            try:
                link.click(no_wait_after=True)
                clicked += 1
                clicked_total += 1
                page.wait_for_timeout(CLICK_DELAY_MS)
            except Exception:
                # fallback: dispatch click
                try:
                    link.evaluate("""(el)=>el.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true,view:window}))""")
                    clicked += 1
                    clicked_total += 1
                    page.wait_for_timeout(CLICK_DELAY_MS)
                except Exception as e2:
                    print(f"  ! Falha ao clicar num recibo: {e2}")

        # aguarda chegarem os PDFs do lote (ou parte deles)
        end = time.time() + BATCH_TIMEOUT_S
        while time.time() < end:
            if len(collector.saved) >= prev + max(1, clicked):
                break
            page.wait_for_timeout(200)

        got = len(collector.saved) - prev
        print(f"  - Lote {base//batch_size + 1}: clicados={clicked} PDFs_salvos={got}")

    # remove listener (evita duplicar se você chamar de novo)
    try:
        ctx.remove_listener("response", collector.on_response)
    except Exception:
        pass

    novos = collector.saved[start_saved:]
    return len(novos), novos

def baixar_recibos_fallback_cdp(page, download_dir: str, max_por_lote: int = None):
    """
    Fallback (Selenium-like): CDP allow download + polling pasta.
    """
    set_download_dir_cdp(page, download_dir)
    links = coletar_links_recibo_da_pagina(page)
    if not links:
        return 0, []
    if max_por_lote is not None:
        links = links[:max_por_lote]

    before = stable_files(download_dir)

    disparados = 0
    for link in links:
        try:
            link.click(no_wait_after=True)
            disparados += 1
            page.wait_for_timeout(CLICK_DELAY_MS)
        except Exception:
            pass

    if disparados == 0:
        return 0, []

    novos = wait_new_files(download_dir, before, expected_min=1, timeout_s=FALLBACK_TIMEOUT_S, quiet_checks=2)
    # pode ter mais de 1, mas devolve todos os novos vistos
    return len(novos), [str(Path(download_dir)/n) for n in novos]

def main():
    Path(DOWNLOAD_DIR).mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(CHROME_DEBUG_ADDR)

        page = selecionar_aba_com_tabela(browser)
        print("Usando aba:", (page.title() or "").strip(), page.url)

        # Opcional: se você quiser forçar o filtro antes
        # page.locator(SEL_FILTRAR).first.click(no_wait_after=True)

        # Método rápido (responses)
        n_fast, paths_fast = baixar_recibos_multi_fast(page, DOWNLOAD_DIR, batch_size=BATCH_SIZE)
        print(f"[FAST] PDFs salvos: {n_fast}")

        # Se veio muito pouco, roda fallback CDP (não fecha o fast; apenas tenta completar)
        if n_fast == 0:
            print("[FAST] Não capturou PDFs via response. Rodando fallback CDP...")
            n_fb, paths_fb = baixar_recibos_fallback_cdp(page, DOWNLOAD_DIR)
            print(f"[FALLBACK] novos arquivos detectados: {n_fb}")

        if paths_fast:
            print("Arquivos (FAST):")
            for pth in paths_fast[:10]:
                print(" -", pth)
            if len(paths_fast) > 10:
                print(f" ... +{len(paths_fast)-10}")

if __name__ == "__main__":
    main()
