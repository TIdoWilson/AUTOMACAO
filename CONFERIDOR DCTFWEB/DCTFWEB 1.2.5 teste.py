import os
import time
from pathlib import Path
from urllib.parse import unquote
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

# Seletor da tabela (igual ao usado no seu script)
SEL_TABELA = "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"

def set_download_dir_cdp(page, download_dir: str):
    """
    Replica o Page.setDownloadBehavior do Selenium: deixa o Chrome salvar
    qualquer download diretamente na pasta informada.
    """
    Path(download_dir).mkdir(parents=True, exist_ok=True)
    try:
        client = page.context.new_cdp_session(page)
        # Page.setDownloadBehavior é aceito pelo Chrome quando via sessão CDP de Page
        client.send("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": download_dir
        })
    except Exception as e:
        # fallback: tenta Browser.setDownloadBehavior (nem sempre disponível)
        try:
            client = page.context.new_cdp_session(page)
            client.send("Browser.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": download_dir
            })
        except Exception as e2:
            raise RuntimeError(f"Não foi possível configurar diretório de download via CDP: {e} / {e2}")

def _stable_files(path_dir: str):
    try:
        allf = os.listdir(path_dir)
    except FileNotFoundError:
        return set()
    # ignora temporários do Chrome
    return {f for f in allf if not f.endswith(".crdownload") and not f.endswith(".tmp") and not f.endswith(".partial")}

def _has_partials(path_dir: str):
    try:
        return any(n.endswith(".crdownload") or n.endswith(".tmp") or n.endswith(".partial")
                   for n in os.listdir(path_dir))
    except FileNotFoundError:
        return False

def _wait_new_files(path_dir: str, before_set: set, expected_min: int, timeout_s: int = 120, quiet_checks: int = 3):
    """
    Espera até que surjam pelo menos 'expected_min' arquivos NOVOS estáveis
    (sem .crdownload) e a pasta fique estável por 'quiet_checks' checagens.
    """
    end = time.time() + timeout_s
    quiet_ok = 0
    while time.time() < end:
        now = _stable_files(path_dir)
        new = now - before_set
        if len(new) >= expected_min and not _has_partials(path_dir):
            quiet_ok += 1
            if quiet_ok >= quiet_checks:
                return list(new)
        else:
            quiet_ok = 0
        time.sleep(0.3)
    raise PWTimeoutError(f"Timeout esperando {expected_min} novos arquivos em '{path_dir}'")

def _all_frames(page):
    return page.frames

def _wait_visible_any_frame(page, selector: str, timeout_ms=60000):
    end = time.time() + (timeout_ms / 1000)
    last = None
    while time.time() < end:
        for fr in _all_frames(page):
            try:
                loc = fr.locator(selector).first
                if loc.count() > 0:
                    loc.wait_for(state="visible", timeout=500)
                    return fr, loc
            except Exception as e:
                last = e
        time.sleep(0.1)
    raise PWTimeoutError(f"Timeout aguardando: {selector}. Último erro: {last}")

def coletar_links_recibo_da_pagina(page):
    """
    Devolve uma lista de Locators (no frame correto) para 'Visualizar Recibo' da página atual.
    Não faz filtragem por Excel; objetivo é só baixar.
    """
    fr, tbl = _wait_visible_any_frame(page, SEL_TABELA, timeout_ms=60000)
    rows = tbl.locator("tr")
    n = rows.count()
    links = []
    for i in range(1, n):  # pula cabeçalho
        tr = rows.nth(i)
        link = tr.locator(
            "a[title='Visualizar Recibo'], a[id*='lkbVisualizarRecibo'], a.image-tabela-visualizar-recibo"
        ).first
        try:
            if link.count() > 0:
                # garante que o locator pertence ao frame da tabela
                links.append(link)
        except Exception:
            pass
    return links

def baixar_recibos_multi(page, download_dir: str, max_por_lote: int = None, atraso_entre_cliques_ms: int = 120):
    """
    Clica rapidamente em vários 'Visualizar Recibo' e espera os arquivos
    aparecerem na pasta (como no Selenium 1.2.3).
    - max_por_lote=None → clica todos da página
    """
    set_download_dir_cdp(page, download_dir)

    # Coleta os links desta página
    links = coletar_links_recibo_da_pagina(page)
    if not links:
        print("Nenhum link de recibo encontrado na página.")
        return 0

    if max_por_lote is not None:
        links = links[:max_por_lote]

    # Snapshot de arquivos antes de disparar
    before = _stable_files(download_dir)

    # Dispara todos os cliques rapidamente (paraleliza os downloads)
    disparados = 0
    for link in links:
        try:
            link.click()
            disparados += 1
            page.wait_for_timeout(atraso_entre_cliques_ms)
        except Exception:
            # fallback: evento de clique via JS no mesmo elemento
            try:
                page.evaluate(
                    """(el)=>{
                        el.dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true, view:window}));
                    }""",
                    link,
                )
                disparados += 1
                page.wait_for_timeout(atraso_entre_cliques_ms)
            except Exception as e2:
                print(f"Falha ao clicar em um recibo: {e2}")

    if disparados == 0:
        print("Nenhum clique de recibo foi disparado.")
        return 0

    # Aguarda chegarem pelo menos 'disparados' novos arquivos
    novos = _wait_new_files(download_dir, before, expected_min=disparados, timeout_s=180, quiet_checks=3)
    print(f"Downloads concluídos: {len(novos)} de {disparados}")
    return len(novos)

def selecionar_aba_ecac(browser):
    """
    Escolhe a aba correta mesmo quando:
    - URL principal não contém ecac
    - ecac está em iframe
    - título é eCAC
    """
    candidatos = []

    for ctx in browser.contexts:
        for p in ctx.pages:
            try:
                url = p.url or ""
            except Exception:
                url = ""
            try:
                title = (p.title() or "").strip()
            except Exception:
                title = ""
            frame_urls = []
            try:
                frame_urls = [(fr.url or "") for fr in p.frames]
            except Exception:
                pass

            candidatos.append((p, title, url, frame_urls))

    if not candidatos:
        raise RuntimeError(
            "Conectou no CDP, mas não há abas visíveis. "
            "Provável Chrome/porta errados."
        )

    # Debug útil (deixe ligado até estabilizar)
    print("Abas encontradas via CDP:")
    for i, (_p, title, url, frame_urls) in enumerate(candidatos):
        print(f"[{i:02d}] title='{title}' url='{url}'")
        for fu in frame_urls:
            if "ecac" in (fu or "").lower():
                print(f"     frame(ecac): {fu}")

    # 1) Melhor caso: algum frame ou a própria aba contém ecac
    for p, title, url, frame_urls in candidatos:
        haystack = " ".join([title, url, *frame_urls]).lower()
        if "ecac" in haystack:
            return p

    # 2) Segundo melhor: aba do eCAC (de onde você entra na ecac)
    for p, title, url, frame_urls in candidatos:
        if "ecac" in (title or "").lower():
            return p

    # 3) Fallback: última aba do último contexto
    return candidatos[-1][0]


DOWNLOAD_DIR = r"T:\testes"

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    page = selecionar_aba_ecac(browser)
    print("Usando aba:", page.title(), page.url)


    # se necessário, clique em 'Filtrar' aqui antes
    # page.locator("a#ctl00_cphConteudo_imageFiltrar, a[id*='imageFiltrar']").first.click()

    # dispara downloads em paralelo (todos da página)
    baixar_recibos_multi(page, DOWNLOAD_DIR, max_por_lote=None)
