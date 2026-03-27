# -*- coding: utf-8 -*-
"""
DCTFWEB 1.2.7 (teste_batch_polling) â€” Playwright + CDP attach (Chrome 9222)
Meta: manter a lÃ³gica do Selenium (baixar vÃ¡rios "ao mesmo tempo") usando:
- CDP: Browser.setDownloadBehavior / Page.setDownloadBehavior para forÃ§ar pasta
- Clique em lote (no_wait_after=True)
- Polling da pasta para detectar novos arquivos (sem expect_download)
- Renomear GUIDs (sem extensÃ£o) para nomes .pdf mais Ãºteis

ObservaÃ§Ãµes de performance:
- Se DOWNLOAD_DIR for unidade de rede (T:, W:), pode ficar bem mais lento.
  Para testar velocidade, use uma pasta local (ex.: C:\\temp\\dctf) e depois mova.
- A sincronizacao dos cliques e guiada por img.image-processamento.

Requisitos:
  pip install playwright
  playwright install chromium
"""

import os
import re
import time
import unicodedata
import tempfile
import subprocess
import urllib.request
from urllib.parse import urlparse
from pathlib import Path
from datetime import datetime
from shutil import which
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
from chrome_9222 import chrome_9222, PORT

# ===================== CONFIG =====================
CHROME_DEBUG_ADDR = os.environ.get("CHROME_DEBUG_ADDR", "http://localhost:9222")
DOWNLOAD_DIR = os.environ.get("DCTF_DOWNLOAD_DIR", r"T:\testes")
ECAC_URL = os.environ.get("DCTF_ECAC_URL", "https://cav.receita.fazenda.gov.br/ecac/")

SEL_TABELA = "table#ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"
SEL_LINK_RECIBO = "a[title='Visualizar Recibo'], a[id*='lkbVisualizarRecibo'], a.image-tabela-visualizar-recibo"
SEL_IMG_PROCESSANDO = "img.image-processamento"

# Ajuste fino
BATCH_SIZE = int(os.environ.get("DCTF_BATCH_SIZE", "10"))          # legado (fluxo atual e sequencial)
BATCH_TIMEOUT_S = int(os.environ.get("DCTF_BATCH_TIMEOUT_S", "60"))# espera novos arquivos
QUIET_CHECKS = int(os.environ.get("DCTF_QUIET_CHECKS", "3"))       # estabilidade da pasta
PROCESSANDO_MAX_WAIT_S = int(os.environ.get("DCTF_PROCESSANDO_MAX_WAIT_S", "5"))
RECIBO_MAX_TENTATIVAS = int(os.environ.get("DCTF_RECIBO_MAX_TENTATIVAS", "2"))
# ==================================================

PARTIAL_SUFFIXES = (".crdownload", ".tmp", ".partial")
RE_INVALID_NAME = re.compile(r"[^\w\-\. ]+", flags=re.UNICODE)
RE_SPACES = re.compile(r"\s+")
RE_REMOTE_DEBUG_PORT = re.compile(r"--remote-debugging-port=(\d+)", flags=re.IGNORECASE)

def log(msg: str):
    print(msg, flush=True)

def find_chrome_exe():
    candidates = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
        which("chrome"),
        which("chrome.exe"),
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    raise FileNotFoundError("chrome.exe nao encontrado. Ajuste o caminho manualmente.")

def parse_debug_addr(addr: str):
    parsed = urlparse(addr if "://" in addr else f"http://{addr}")
    host = parsed.hostname or "127.0.0.1"
    port = parsed.port or 9222
    return host, port

def cdp_ready(addr: str):
    host, port = parse_debug_addr(addr)
    url = f"http://{host}:{port}/json/version"
    try:
        with urllib.request.urlopen(url, timeout=1) as r:
            return r.status == 200
    except Exception:
        return False

def chrome_running() -> bool:
    if os.name != "nt":
        return False
    try:
        p = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=3
        )
        out = (p.stdout or "").lower()
        if "chrome.exe" not in out:
            return False
        if "no tasks are running" in out or "nenhuma tarefa" in out:
            return False
        return True
    except Exception:
        return False

def detect_running_chrome_debug_ports() -> list[int]:
    if os.name != "nt":
        return []

    cmd = [
        "powershell",
        "-NoProfile",
        "-Command",
        "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | Select-Object -ExpandProperty CommandLine"
    ]
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
        if p.returncode != 0:
            return []
        ports = []
        for line in (p.stdout or "").splitlines():
            m = RE_REMOTE_DEBUG_PORT.search(line or "")
            if not m:
                continue
            try:
                ports.append(int(m.group(1)))
            except Exception:
                continue
        return sorted(set(ports))
    except Exception:
        return []

def resolver_debug_addr(addr_preferido: str):
    """
    Retorna (addr, origem):
      - origem = "preferido" quando o endereco configurado ja responde CDP
      - origem = "detectado" quando encontrou outra porta CDP em Chrome ja aberto
      - origem = "chrome_sem_cdp" quando Chrome existe sem CDP ativo
      - origem = "inexistente" quando nao encontrou Chrome aberto
    """
    if cdp_ready(addr_preferido):
        return addr_preferido, "preferido"

    host_pref, port_pref = parse_debug_addr(addr_preferido)
    candidate_ports = [port_pref]
    for p in detect_running_chrome_debug_ports():
        if p not in candidate_ports:
            candidate_ports.append(p)

    for p in candidate_ports:
        probe = f"http://127.0.0.1:{p}"
        if cdp_ready(probe):
            return probe, "detectado"

    if chrome_running():
        return addr_preferido, "chrome_sem_cdp"
    return addr_preferido, "inexistente"

def launch_chrome_with_cdp(addr: str, start_url: str):
    host, port = parse_debug_addr(addr)
    chrome = find_chrome_exe()
    user_data = str(Path(tempfile.gettempdir()) / f"chrome-dctfweb-{port}")
    os.makedirs(user_data, exist_ok=True)
    args = [
        chrome,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={user_data}",
        "--new-window",
        start_url,
    ]
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(60):
        if cdp_ready(addr):
            return
        time.sleep(0.25)
    raise TimeoutError(
        f"Chrome com CDP nao respondeu em {host}:{port}. "
        f"Se houver Chrome aberto sem depuracao, feche todas as janelas do Chrome e tente novamente."
    )

def garantir_chrome_cdp(addr: str, start_url: str):
    if cdp_ready(addr):
        return
    log(f"Chrome com CDP nao encontrado em {addr}. Abrindo navegador...")
    launch_chrome_with_cdp(addr, start_url)

def abrir_ou_focar_ecac(browser, ecac_url: str):
    for ctx in browser.contexts:
        for pg in ctx.pages:
            try:
                if "ecac" in (pg.url or "").lower():
                    pg.bring_to_front()
                    return pg
            except Exception:
                pass

    if browser.contexts:
        ctx = browser.contexts[0]
    else:
        ctx = browser.new_context()
    pg = ctx.new_page()
    pg.goto(ecac_url, wait_until="domcontentloaded")
    pg.bring_to_front()
    return pg

def sanitize_filename(s: str) -> str:
    s = s.strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = RE_INVALID_NAME.sub("_", s)
    s = RE_SPACES.sub("_", s)
    s = s.strip("._")
    return s[:180] if len(s) > 180 else s

def scan_dir_state(path_dir: str):
    """
    Faz uma Ãºnica varredura da pasta e retorna:
      - conjunto de arquivos estÃ¡veis (sem temporÃ¡rios)
      - flag se ainda existem parciais
    """
    try:
        names = [entry.name for entry in os.scandir(path_dir) if entry.is_file()]
    except FileNotFoundError:
        return set(), False

    has_partial = any(name.endswith(PARTIAL_SUFFIXES) for name in names)
    stable = {name for name in names if not name.endswith(PARTIAL_SUFFIXES)}
    return stable, has_partial

def stable_files(path_dir: str):
    stable, _ = scan_dir_state(path_dir)
    return stable

def montar_pastas_monitoradas(download_dir: str):
    pastas = [str(Path(download_dir))]
    fallback = str(Path.home() / "Downloads")
    if fallback not in pastas:
        pastas.append(fallback)

    extra = os.environ.get("DCTF_EXTRA_WATCH_DIRS", "").strip()
    if extra:
        for p in extra.split(";"):
            p = p.strip()
            if p and p not in pastas:
                pastas.append(p)
    return pastas

def wait_new_files(path_dir: str, before_set: set, expected_min: int, timeout_s: int, quiet_checks: int):
    end = time.monotonic() + timeout_s
    quiet_ok = 0
    last_new_count = 0
    while time.monotonic() < end:
        now, partials_present = scan_dir_state(path_dir)
        new = list(now - before_set)
        new_count = len(new)
        if new_count >= expected_min and not partials_present:
            # espera a pasta ficar estÃ¡vel por algumas checagens
            if new_count == last_new_count:
                quiet_ok += 1
            else:
                quiet_ok = 0
                last_new_count = new_count
            if quiet_ok >= quiet_checks:
                return new
        else:
            quiet_ok = 0
            last_new_count = new_count
        time.sleep(0.25)
    raise PWTimeoutError(f"Timeout esperando {expected_min} novos arquivos em '{path_dir}' (vi {last_new_count}).")

def wait_new_files_multi(watch_dirs: list[str], before_map: dict, expected_min: int, timeout_s: int, quiet_checks: int):
    end = time.monotonic() + timeout_s
    quiet_ok = 0
    last_new_count = 0

    while time.monotonic() < end:
        new_entries = []
        partials_present = False

        for d in watch_dirs:
            now, has_partial = scan_dir_state(d)
            partials_present = partials_present or has_partial
            before = before_map.get(d, set())
            for fname in (now - before):
                new_entries.append((d, fname))

        new_count = len(new_entries)
        if new_count >= expected_min and not partials_present:
            if new_count == last_new_count:
                quiet_ok += 1
            else:
                quiet_ok = 0
                last_new_count = new_count
            if quiet_ok >= quiet_checks:
                return new_entries
        else:
            quiet_ok = 0
            last_new_count = new_count

        time.sleep(0.25)

    raise PWTimeoutError(f"Timeout esperando {expected_min} novos arquivos nas pastas monitoradas (vi {last_new_count}).")

def wait_until_quiet_multi(
    watch_dirs: list[str],
    before_map: dict,
    timeout_s: int,
    quiet_checks: int,
    expected_min: int = 0
):
    """
    Aguarda a fila de download estabilizar e retorna todos os novos arquivos
    detectados nas pastas monitoradas.
    """
    end = time.monotonic() + timeout_s
    quiet_ok = 0
    last_new_count = -1
    last_entries = []

    while time.monotonic() < end:
        entries = []
        partials_present = False
        for d in watch_dirs:
            now, has_partial = scan_dir_state(d)
            partials_present = partials_present or has_partial
            before = before_map.get(d, set())
            for fname in (now - before):
                entries.append((d, fname))

        new_count = len(entries)
        enough_files = new_count >= expected_min
        if not partials_present and enough_files:
            if new_count == last_new_count:
                quiet_ok += 1
            else:
                quiet_ok = 0
                last_new_count = new_count
                last_entries = entries

            if quiet_ok >= quiet_checks:
                return last_entries
        else:
            quiet_ok = 0
            last_new_count = new_count
            last_entries = entries

        time.sleep(0.25)

    return last_entries

def wait_visible_any_frame(page, selector: str, timeout_ms=60000):
    end = time.monotonic() + (timeout_ms / 1000)
    last = None
    while time.monotonic() < end:
        for fr in page.frames:
            try:
                el = fr.query_selector(selector)
                if el and el.is_visible():
                    return fr, fr.locator(selector).first
            except Exception as e:
                last = e
        time.sleep(0.1)
    raise PWTimeoutError(f"Timeout aguardando: {selector}. Ãšltimo erro: {last}")

def imagem_processando_visivel(page) -> bool:
    for fr in page.frames:
        try:
            loc = fr.locator(SEL_IMG_PROCESSANDO).first
            if loc.count() > 0 and loc.is_visible():
                return True
        except Exception:
            continue
    return False

def wait_processamento_sumir(page, timeout_ms=60000):
    end = time.monotonic() + (timeout_ms / 1000)
    while time.monotonic() < end:
        if not imagem_processando_visivel(page):
            return
        time.sleep(0.15)
    raise PWTimeoutError(f"Timeout aguardando sumir: {SEL_IMG_PROCESSANDO}")

def clicar_imagem_processamento(page) -> bool:
    for fr in page.frames:
        try:
            loc = fr.locator(SEL_IMG_PROCESSANDO).first
            if loc.count() > 0 and loc.is_visible():
                loc.click(no_wait_after=True, timeout=1500, force=True)
                return True
        except Exception:
            continue
    return False

def garantir_processamento_livre(page, contexto: str = "", timeout_s: int = PROCESSANDO_MAX_WAIT_S) -> bool:
    # Aguarda apenas 2/3 do timeout para tentar destravar mais cedo.
    timeout_pre_click_s = max(1, int(round(max(1, timeout_s) * (2.0 / 3.0))))
    try:
        wait_processamento_sumir(page, timeout_ms=timeout_pre_click_s * 1000)
        return True
    except Exception:
        if not clicar_imagem_processamento(page):
            return False
        if contexto:
            log(f"  - Processando persistiu ({contexto}), cliquei na imagem para destravar.")
        try:
            wait_processamento_sumir(page, timeout_ms=max(1, timeout_s) * 1000)
            return True
        except Exception:
            return False

def selecionar_aba_com_tabela(browser):
    pages = []
    for ctx in browser.contexts:
        pages.extend(ctx.pages)

    if not pages:
        raise RuntimeError("Conectou no CDP mas nÃ£o hÃ¡ abas visÃ­veis. Verifique a instÃ¢ncia/porta 9222.")

    log("Abas vistas via CDP:")
    for i, pg in enumerate(pages):
        try:
            title = (pg.title() or "").strip()
        except Exception:
            title = ""
        try:
            url = pg.url or ""
        except Exception:
            url = ""
        log(f"[{i:02d}] title='{title}' url='{url}'")

    # preferir onde a tabela existe em algum frame
    for pg in pages:
        try:
            for fr in pg.frames:
                try:
                    if fr.query_selector(SEL_TABELA):
                        return pg
                except Exception:
                    pass
        except Exception:
            pass

    # fallback: algo que pareÃ§a dctf/ecac
    for pg in pages:
        hay = ""
        try: hay += (pg.url or "")
        except Exception: pass
        try: hay += " " + (pg.title() or "")
        except Exception: pass
        if "dctf" in hay.lower() or "ecac" in hay.lower():
            return pg

    return pages[-1]

def set_download_behavior_global(page, download_dir: str):
    """
    Tenta aplicar globalmente (Browser) e no target atual (Page), para pegar downloads
    disparados por diferentes caminhos.
    """
    Path(download_dir).mkdir(parents=True, exist_ok=True)
    sess = page.context.new_cdp_session(page)

    # Browser.* (mais "global")
    try:
        sess.send("Browser.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": download_dir,
            "eventsEnabled": True
        })
        log("CDP: Browser.setDownloadBehavior OK")
    except Exception as e:
        log(f"CDP: Browser.setDownloadBehavior falhou: {e}")

    # Page.* (por garantia)
    try:
        sess.send("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": download_dir
        })
        log("CDP: Page.setDownloadBehavior OK")
    except Exception as e:
        log(f"CDP: Page.setDownloadBehavior falhou: {e}")

def extrair_info_linha(tr):
    """
    ExtraÃ§Ã£o rÃ¡pida via evaluate (bem mais estÃ¡vel que vÃ¡rios inner_text).
    Retorna dict com campos Ãºteis para nomear arquivo.
    """
    cells = tr.evaluate("""(row) => Array.from(row.querySelectorAll('td')).map(td => (td.innerText||'').trim())""")
    numero = cells[1] if len(cells) > 1 else ""
    periodo = cells[2] if len(cells) > 2 else ""
    data_tx = cells[3] if len(cells) > 3 else ""
    return {
        "numero": numero,
        "periodo": periodo,
        "data": data_tx,
        "cells": cells,
    }

def fechar_loading_pos_clique(fr, link):
    """
    ApÃ³s clicar no recibo, tenta desfocar o link e clicar fora para fechar o
    "processando" antes do prÃ³ximo clique.
    """
    try:
        link.evaluate("el => el.blur()")
    except Exception:
        pass

    try:
        fr.locator("body").click(position={"x": 5, "y": 5}, no_wait_after=True, timeout=1500)
    except Exception:
        pass

def coletar_links_e_nomes(page):
    fr, tbl = wait_visible_any_frame(page, SEL_TABELA, timeout_ms=60000)
    rows = tbl.locator("tr")
    n = rows.count()
    items = []
    for i in range(1, n):
        tr = rows.nth(i)
        link = tr.locator(SEL_LINK_RECIBO).first
        if link.count() == 0:
            continue
        info = extrair_info_linha(tr)
        base = f"{info['numero']}_{info['periodo']}_{info['data']}"
        base = sanitize_filename(base) or f"recibo_linha_{i:03d}"
        desired = f"{base}.pdf"
        items.append((link, desired, i, fr))
    return items

def assinatura_tabela(tbl):
    try:
        return tbl.evaluate(
            """(table) => {
                const rows = Array.from(table.querySelectorAll('tr')).slice(1, 6);
                const sigRows = rows.map((r) => {
                    const t = (r.innerText || '').replace(/\\s+/g, ' ').trim();
                    return t.slice(0, 120);
                });
                return `${rows.length}|${sigRows.join('||')}`;
            }"""
        )
    except Exception:
        return ""

def ir_para_proxima_pagina(page, timeout_ms=30000):
    if not garantir_processamento_livre(page, contexto="antes da troca de pagina"):
        return False
    fr, _ = wait_visible_any_frame(page, SEL_TABELA, timeout_ms=60000)

    candidatos = [
        "a[href*='Page$Next']",
        "a[href*='Page%24Next']",
        "a[href*='__doPostBack'][href*='Next']",
        "a[title*='Proxima']",
        "a[id*='Proxima']",
        "a[id*='Proximo']",
        "a[id*='lnkProx']",
        "a[id*='lkbProx']",
        "a[aria-label*='Proxima']",
        "a:has-text('Proxima')",
        "a:has-text('>')",
    ]

    botao = None
    for sel in candidatos:
        loc = fr.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible():
                css = (loc.get_attribute("class") or "").lower()
                if "disabled" in css:
                    continue
                botao = loc
                break
        except Exception:
            continue

    if botao is None:
        return False

    try:
        botao.click(no_wait_after=True)
    except Exception:
        return False

    if not garantir_processamento_livre(page, contexto="apos troca de pagina"):
        return False
    wait_visible_any_frame(page, SEL_TABELA, timeout_ms=timeout_ms)
    return True
def sort_by_mtime(dirpath: str, filenames: list[str]) -> list[str]:
    if not filenames:
        return []

    wanted = set(filenames)
    mtimes = {}
    try:
        for entry in os.scandir(dirpath):
            if not entry.is_file():
                continue
            if entry.name in wanted:
                try:
                    mtimes[entry.name] = entry.stat().st_mtime
                except FileNotFoundError:
                    mtimes[entry.name] = 0.0
    except FileNotFoundError:
        return filenames

    return sorted(filenames, key=lambda name: mtimes.get(name, 0.0))

def sort_entries_by_mtime(entries: list[tuple[str, str]]) -> list[tuple[str, str]]:
    enriched = []
    for d, f in entries:
        p = Path(d) / f
        try:
            mt = p.stat().st_mtime
        except Exception:
            mt = 0.0
        enriched.append((mt, d, f))
    enriched.sort(key=lambda x: x[0])
    return [(d, f) for _, d, f in enriched]

def rename_to_pdf(dirpath: str, src_name: str, dst_name: str, target_dir: str | None = None):
    src = Path(dirpath) / src_name
    if not src.exists():
        return None

    # se nÃ£o tem extensÃ£o, mas parece PDF, forÃ§a .pdf
    if src.suffix == "":
        try:
            with src.open("rb") as f:
                if f.read(5) == b"%PDF-" and not dst_name.lower().endswith(".pdf"):
                    dst_name = dst_name + ".pdf"
        except Exception:
            pass

    dst_base = Path(target_dir) if target_dir else Path(dirpath)
    dst_base.mkdir(parents=True, exist_ok=True)
    dst = dst_base / dst_name
    if dst.exists():
        stem, suf = dst.stem, dst.suffix
        dst = Path(dirpath) / f"{stem}_{datetime.now().strftime('%H%M%S%f')}{suf}"

    try:
        src.rename(dst)
        return str(dst)
    except Exception:
        try:
            data = src.read_bytes()
            dst.write_bytes(data)
            try:
                src.unlink()
            except Exception:
                pass
            return str(dst)
        except Exception:
            return None

def baixar_recibos_em_lotes(page, download_dir: str, pagina_idx: int = 1):
    Path(download_dir).mkdir(parents=True, exist_ok=True)
    set_download_behavior_global(page, download_dir)
    watch_dirs = montar_pastas_monitoradas(download_dir)
    log("Pastas monitoradas: " + " | ".join(watch_dirs))

    items = coletar_links_e_nomes(page)
    if not items:
        log(f"Pagina {pagina_idx}: nenhum recibo encontrado.")
        return 0, 0, 0

    log(f"Pagina {pagina_idx}: recibos detectados = {len(items)}")

    total_ok = 0
    total_fail = 0
    timeout_por_recibo = max(BATCH_TIMEOUT_S, 30)

    # Processamento sequencial com retry do mesmo recibo quando nao baixa.
    max_tentativas = max(1, RECIBO_MAX_TENTATIVAS)
    for idx, (link, desired, linha_idx, _fr) in enumerate(items, start=1):
        sucesso = False
        ultimo_erro = ""

        for tentativa in range(1, max_tentativas + 1):
            if not garantir_processamento_livre(page, contexto=f"antes linha {linha_idx}", timeout_s=PROCESSANDO_MAX_WAIT_S):
                ultimo_erro = "processando nao encerrou antes do clique"
                if tentativa < max_tentativas:
                    log(f"  - {idx:02d}/{len(items)} linha {linha_idx}: retry {tentativa+1}/{max_tentativas} apos travamento de processando.")
                continue

            before_click = {d: stable_files(d) for d in watch_dirs}
            try:
                link.click(no_wait_after=True)
            except Exception as e:
                ultimo_erro = f"falha no clique: {e}"
                if tentativa < max_tentativas:
                    log(f"  - {idx:02d}/{len(items)} linha {linha_idx}: retry {tentativa+1}/{max_tentativas} apos erro de clique.")
                continue

            processamento_livre_pos = garantir_processamento_livre(
                page,
                contexto=f"apos linha {linha_idx}",
                timeout_s=PROCESSANDO_MAX_WAIT_S
            )
            if not processamento_livre_pos:
                ultimo_erro = "processando nao encerrou apos clique"

            try:
                new_entries = wait_new_files_multi(
                    watch_dirs=watch_dirs,
                    before_map=before_click,
                    expected_min=1,
                    timeout_s=timeout_por_recibo,
                    quiet_checks=max(QUIET_CHECKS, 2),
                )
                new_entries_sorted = sort_entries_by_mtime(new_entries)
                src_dir, src_name = new_entries_sorted[0]
                out = rename_to_pdf(src_dir, src_name, desired, target_dir=download_dir)
                if not out:
                    ultimo_erro = "falha ao renomear arquivo baixado"
                    if tentativa < max_tentativas:
                        log(f"  - {idx:02d}/{len(items)} linha {linha_idx}: retry {tentativa+1}/{max_tentativas} apos falha de renomeacao.")
                    continue

                total_ok += 1
                sucesso = True
                log(f"  - {idx:02d}/{len(items)} OK linha {linha_idx}: {Path(out).name}")
                break
            except Exception as e:
                if not processamento_livre_pos:
                    ultimo_erro = f"processando travado e nenhum download detectado: {e}"
                else:
                    ultimo_erro = f"nenhum download detectado: {e}"
                if tentativa < max_tentativas:
                    log(f"  - {idx:02d}/{len(items)} linha {linha_idx}: sem download, repetindo ({tentativa+1}/{max_tentativas}).")

        if not sucesso:
            total_fail += 1
            log(f"  ! {idx:02d}/{len(items)} falha linha {linha_idx}: {ultimo_erro}")

    log(f"Pagina {pagina_idx}: renomeados OK = {total_ok} | Falhas/pendentes = {total_fail}")
    return total_ok, total_fail, len(items)

def baixar_todas_paginas(page, download_dir: str):
    pagina = 1
    total_geral_ok = 0
    total_geral_fail = 0
    total_recibos = 0
    assinaturas_vistas = set()

    while True:
        try:
            _, tbl = wait_visible_any_frame(page, SEL_TABELA, timeout_ms=60000)
            sig = assinatura_tabela(tbl)
            if sig in assinaturas_vistas:
                log("Paginacao interrompida: tabela repetida (sem mudanca de pagina).")
                break
            assinaturas_vistas.add(sig)
        except Exception:
            pass

        ok, fail, qtd = baixar_recibos_em_lotes(page, download_dir, pagina_idx=pagina)
        total_geral_ok += ok
        total_geral_fail += fail
        total_recibos += qtd

        avancou = ir_para_proxima_pagina(page)
        if not avancou:
            break

        pagina += 1
        log(f"Avancando para pagina {pagina}...")

    log(f"Concluido geral. Paginas processadas: {pagina}")
    log(f"Total de recibos detectados: {total_recibos}")
    log(f"Total renomeados OK: {total_geral_ok} | Falhas/pendentes: {total_geral_fail}")
    log(f"Pasta: {download_dir}")
def main():
    Path(DOWNLOAD_DIR).mkdir(parents=True, exist_ok=True)
    log(f"Preparando Chrome em {CHROME_DEBUG_ADDR}")
    log(f"DOWNLOAD_DIR = {DOWNLOAD_DIR}")

    debug_addr, origem_addr = resolver_debug_addr(CHROME_DEBUG_ADDR)
    host, port = parse_debug_addr(debug_addr)
    if host not in ("127.0.0.1", "localhost"):
        log(f"CHROME_DEBUG_ADDR com host '{host}' nao e suportado no helper chrome_9222; usando localhost:{port}.")
        debug_addr = f"http://127.0.0.1:{port}"
        host, port = parse_debug_addr(debug_addr)

    if origem_addr == "preferido":
        log(f"Navegador de conexao ja esta aberto em {debug_addr}.")
    elif origem_addr == "detectado":
        log(f"Chrome ja aberto detectado com CDP em {debug_addr}. Reutilizando essa instancia.")
    elif origem_addr == "chrome_sem_cdp":
        log("Chrome ja esta aberto, mas sem CDP ativo. Abrindo instancia controlavel para automacao...")
        launch_chrome_with_cdp(debug_addr, ECAC_URL)
    else:
        log("Nenhum Chrome com CDP detectado. Abrindo navegador para login no eCAC...")
        launch_chrome_with_cdp(debug_addr, ECAC_URL)

    log("Faca login manual no eCAC e deixe a tela da DCTFWeb pronta.")
    try:
        resp = input("Quando terminar o login, digite S para conectar e executar [s/N]: ").strip().lower()
    except EOFError:
        resp = ""
    if resp not in ("s", "sim"):
        log("Operacao cancelada antes da conexao com o navegador.")
        return

    with sync_playwright() as p:
        browser = chrome_9222(p, port=port or PORT)
        page = selecionar_aba_com_tabela(browser)
        log(f"Usando aba: {(page.title() or '').strip()} | {page.url}")
        baixar_todas_paginas(page, DOWNLOAD_DIR)
        browser.close()

if __name__ == "__main__":
    main()


