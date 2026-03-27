#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script: eco_class_ncm_scraper.py
Versão: 5 (extração estruturada sem cabeçalhos)
- Mantém logs detalhados e recuperação robusta.
- Extrai itens das abas "CST", "cClassTrib", "cCredPres" em formato "COD - Descrição" sem os cabeçalhos.
- Fallback por texto se o DOM não tiver a estrutura esperada.
- Continua salvando o Excel final na mesma pasta do Excel selecionado.
"""

import os
import re
import sys
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple

# ---- Dependências ----
try:
    import pandas as pd
except ImportError:
    print(">> Você precisa instalar pandas:  pip install pandas")
    sys.exit(1)

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except ImportError:
    print(">> Você precisa instalar playwright:  pip install playwright && playwright install")
    sys.exit(1)

# ---- Configurações ----
CDP_ENDPOINT = "http://localhost:9222"
FINAL_LINK = "https://app.econeteditora.com.br/app/eco-class"

# Delays/Timeouts
SEARCH_TYPING_DELAY_MS = 15
POST_SEARCH_SLEEP = 0.15
SMALL_WAIT_MS = 120
TAB_WAIT_MS = 180
ROW_OPEN_WAIT_MS = 240

DEFAULT_WAIT_TABLE_MS = 3000
DEFAULT_WAIT_HEADERS_MS = 3500
DEFAULT_WAIT_BADGE_RETRY_MS = 700
DEFAULT_BADGE_RETRIES = 8

# Seletores candidatos para o input de busca (id varia)
SEARCH_INPUT_CANDIDATES = [
    'input#input-v-34',
    'input.v-field__input[type="text"]',
    'input[type="text"][aria-describedby*="input-v-"]',
    'input[type="text"][placeholder*="NCM" i]',
    'input[type="text"]'
]

# Rótulos das abas
DETAIL_TABS = {
    "CST": ["CST", "cst"],
    "cClassTrib": ["cClassTrib", "cclasstrib", "cclass", "class", "class trib"],
    "cCredPres": ["cCredPres", "ccredpres", "cred", "credito", "c cred"],
    "Alíquota": ["Alíquota", "Aliquota", "aliquota", "alíquota"]
}

# ------------------------- LOGGING -------------------------
def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log(msg: str) -> None:
    print(f"[{now()}] {msg}", flush=True)

def log_step(step: str) -> float:
    log(step + " ...")
    return time.perf_counter()

def log_done(start: float, ok_msg: str) -> None:
    dt = (time.perf_counter() - start) * 1000.0
    log(f"{ok_msg} (em {dt:.0f} ms)")

def log_warn(msg: str) -> None:
    print(f"[{now()}] [AVISO] {msg}", flush=True)

def log_info(msg: str) -> None:
    print(f"[{now()}] [INFO] {msg}", flush=True)

def log_err(msg: str) -> None:
    print(f"[{now()}] [ERRO] {msg}", flush=True)


# ---------------------- UTILITÁRIOS ------------------------
def normalize_header(s: str) -> str:
    if s is None:
        return ""
    s = s.strip().lower()
    acentos = str.maketrans("áàâãäéèêëíìîïóòôõöúùûüç", "aaaaaeeeeiiiiooooouuuuc")
    s = s.translate(acentos)
    s = re.sub(r"[\s\W_]+", "", s)
    return s

def detect_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    cols = list(df.columns)
    norm = {c: normalize_header(str(c)) for c in cols}

    ncm_col = None
    for c, n in norm.items():
        if "ncm" in n:
            ncm_col = c
            break

    codigo_alvos = [
        "codigo", "codigodoproduto", "codigoproduto", "codproduto", "codprod",
        "sku", "codigoitem", "cod", "codigoerp"
    ]
    codigo_col = None
    for c, n in norm.items():
        if any(alvo == n for alvo in codigo_alvos):
            codigo_col = c
            break

    nome_alvos = ["nome", "descricao", "descricaoproduto", "nomedoproduto", "produto"]
    nome_col = None
    for c, n in norm.items():
        if any(alvo == n for alvo in nome_alvos):
            nome_col = c
            break

    return codigo_col, nome_col, ncm_col

def only_digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def ncm_to_8digits(ncm: str) -> str:
    n = only_digits(str(ncm))
    return n[:8] if len(n) >= 8 else n

def ask_excel_path() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel com os NCMs",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            return path
    except Exception:
        pass

    print("Informe o caminho completo do arquivo Excel (ex: C:\\dados\\planilha.xlsx):")
    return input("> ").strip('"').strip()


# -------------------- PLAYWRIGHT HELPERS -------------------
def connect_to_chrome_over_cdp(cdp_endpoint: str):
    pw = sync_playwright().start()
    try:
        browser = pw.chromium.connect_over_cdp(cdp_endpoint)
        return pw, browser
    except Exception as e:
        pw.stop()
        raise RuntimeError(
            f"Não consegui conectar ao Chrome via CDP em {cdp_endpoint}.\n"
            f"Abra o Chrome com --remote-debugging-port=9222.\nErro: {e}"
        )

def get_or_open_eco_class_page(browser, final_link: str):
    target_page = None
    for context in browser.contexts:
        for page in context.pages:
            if page.url.startswith(final_link):
                target_page = page
                break
        if target_page:
            break

    if not target_page:
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        target_page = context.new_page()
        target_page.goto(final_link, wait_until="domcontentloaded")
        target_page.wait_for_timeout(SMALL_WAIT_MS)
    try:
        target_page.bring_to_front()
    except Exception:
        pass
    return target_page

def find_search_input(page):
    for sel in SEARCH_INPUT_CANDIDATES:
        try:
            loc = page.locator(sel).first
            if loc and loc.is_visible():
                log_info(f"Input de busca localizado por seletor: {sel}")
                return loc
        except Exception:
            continue
    try:
        loc = page.get_by_role("textbox").first
        if loc.is_visible():
            log_info("Input de busca localizado via role=textbox (fallback).")
            return loc
    except Exception:
        pass
    raise RuntimeError("Não consegui localizar o campo de busca de NCM na página.")

def wait_for_variations_table(page, timeout_ms: int = DEFAULT_WAIT_TABLE_MS) -> bool:
    try:
        page.locator("div.tab-content table tbody tr").first.wait_for(state="visible", timeout=timeout_ms)
        return True
    except PlaywrightTimeoutError:
        return False

def wait_for_header_tabs(page, timeout_ms: int = DEFAULT_WAIT_HEADERS_MS) -> bool:
    try:
        page.locator("div.tab-header-item").first.wait_for(state="visible", timeout=timeout_ms)
        return True
    except PlaywrightTimeoutError:
        return False

def click_tab_by_labels(page, labels: List[str]) -> bool:
    for t in labels:
        try:
            tab = page.locator(f"div.tab-header-item p:has-text('{t}')").first
            if tab and tab.is_visible():
                tab.click()
                page.wait_for_timeout(TAB_WAIT_MS)
                return True
        except Exception:
            continue
    for t in labels:
        try:
            page.get_by_text(t, exact=False).first.click()
            page.wait_for_timeout(TAB_WAIT_MS)
            return True
        except Exception:
            continue
    return False

def clear_econet_cookies_keep_session(page) -> None:
    """
    Limpa apenas cookies 'não essenciais' do domínio econeteditora.com.br,
    tentando preservar cookies de sessão/autenticação.

    Estratégia:
      - Captura todos os cookies do contexto.
      - Mantém:
          * todos os cookies que NÃO são da econeteditora.com.br
          * cookies da econeteditora.com.br cujo nome sugere sessão/autenticação
      - Limpa todos os cookies do contexto.
      - Reinsere apenas os cookies preservados.
    """
    try:
        ctx = page.context
        all_cookies = ctx.cookies()
    except Exception as e:
        log_warn(f"Falha ao ler cookies do contexto: {e}")
        return

    if not all_cookies:
        log_info("Nenhum cookie encontrado para limpar.")
        return

    def is_session_cookie(c: dict) -> bool:
        name = (c.get("name") or "").lower()
        # heurística para cookies de sessão/autenticação
        session_tokens = [
            "sess", "session", "auth", "login",
            "token", "jwt", "asp.net", "jsession"
        ]
        return any(tok in name for tok in session_tokens)

    keep_cookies = []
    removed_cookies = []

    for c in all_cookies:
        domain = (c.get("domain") or "").lower()
        if "econeteditora.com.br" not in domain:
            # cookies de outros domínios sempre mantidos
            keep_cookies.append(c)
        else:
            # cookies da econet: só mantemos se parecer de sessão
            if is_session_cookie(c):
                keep_cookies.append(c)
            else:
                removed_cookies.append(c)

    if not removed_cookies:
        log_info("Nenhum cookie da Econet identificado para limpeza (mantidos como estão).")
        return

    try:
        # zera tudo
        ctx.clear_cookies()
        # reinsere apenas os cookies 'preservados'
        ctx.add_cookies(keep_cookies)
        log_info(
            f"Limpamos {len(removed_cookies)} cookie(s) da econeteditora.com.br "
            f"e preservamos {len(keep_cookies)} cookie(s) (incluindo sessão)."
        )
    except Exception as e:
        log_warn(f"Erro ao limpar/reinserir cookies da Econet: {e}")

# --- Leitura de badge com retries (v6 robusta) ---
def read_badge_with_retries(page, labels: List[str], retries: int = DEFAULT_BADGE_RETRIES, sleep_ms: int = DEFAULT_WAIT_BADGE_RETRY_MS) -> int:
    """
    Lê o número do badge da aba indicada por 'labels' (case-insensitive), com reintentos.
    Estratégias:
      - scroll até o header
      - tenta localizar <span.badge> (ou qualquer [class*=badge])
      - fallback: parseia número no texto do header (ex.: "CST 2")
    Retorna 0 se não estabilizar/achar nada.
    """
    import re
    last_val: Optional[int] = None
    stable_reads = 0

    def parse_int(s: str) -> Optional[int]:
        m = re.search(r"\d+", s or "")
        return int(m.group(0)) if m else None

    for attempt in range(1, retries + 1):
        for label in labels:
            try:
                # localiza o <p> do header, insensível a maiúsculas
                p_loc = page.locator("div.tab-header-item p", has_text=re.compile(label, re.IGNORECASE)).first
                if not p_loc or not p_loc.is_visible():
                    continue

                # garante que está no viewport
                try:
                    p_loc.scroll_into_view_if_needed()
                    page.wait_for_timeout(80)
                except Exception:
                    pass

                # 1) tenta badge padrão
                badge = p_loc.locator("span.badge").first
                try:
                    badge.wait_for(state="visible", timeout=1200)
                    raw = (badge.inner_text() or "").strip()
                    val = parse_int(raw)
                    if val is not None:
                        if last_val is None:
                            last_val, stable_reads = val, 1
                        else:
                            if val == last_val:
                                stable_reads += 1
                                if stable_reads >= 2:
                                    return val
                            else:
                                last_val, stable_reads = val, 1
                        continue  # já leu via badge; próxima tentativa (para estabilizar)
                except Exception:
                    pass

                # 2) seletor alternativo de badge (alguns temas)
                try:
                    alt_badge = p_loc.locator("[class*=badge]").first
                    if alt_badge and alt_badge.is_visible():
                        raw = (alt_badge.inner_text() or "").strip()
                        val = parse_int(raw)
                        if val is not None:
                            if last_val is None:
                                last_val, stable_reads = val, 1
                            else:
                                if val == last_val:
                                    stable_reads += 1
                                    if stable_reads >= 2:
                                        return val
                                else:
                                    last_val, stable_reads = val, 1
                            continue
                except Exception:
                    pass

                # 3) fallback: tenta extrair o número do próprio texto do header
                try:
                    header_text = (p_loc.inner_text() or "").strip()   # ex.: "CST 2"
                    val = parse_int(header_text)
                    if val is not None:
                        if last_val is None:
                            last_val, stable_reads = val, 1
                        else:
                            if val == last_val:
                                stable_reads += 1
                                if stable_reads >= 2:
                                    return val
                            else:
                                last_val, stable_reads = val, 1
                except Exception:
                    pass

            except Exception:
                continue

        time.sleep(sleep_ms / 1000.0)

    return last_val if isinstance(last_val, int) else 0


# ------------------- EXTRAÇÃO E LIMPEZA --------------------
HEADER_TOKENS = {
    "cst", "descrição", "descricao", "tipo de benefício", "tipo de beneficio",
    "ncm", "alíquota", "aliquota", "cclasstrib", "ccredpres", "cclass", "class"
}

def is_header_line(s: str) -> bool:
    t = normalize_header(s)
    return t in HEADER_TOKENS or t == ""

def collapse_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def extract_aliquota_struct(page) -> list[dict]:
    """
    Lê a aba 'Alíquota' de forma estruturada abrindo cada painel (IBS, CBS, ...).
    Retorna lista de dicts:
      {'trib': 'IBS', 'titulo': '...', 'percent': '0,90%',
       'vigencia': '01/01/2026 - 31/12/2026', 'base_legal': 'Artigo ...'}
    """
    import re
    items: list[dict] = []

    try:
        headers = page.locator("div.tab-content .expansion-panel-header")
        cnt = headers.count()

        for i in range(cnt):
            header = headers.nth(i)

            # 0) Garante que o header está visível e na tela
            try:
                header.scroll_into_view_if_needed()
                page.wait_for_timeout(120)
            except Exception:
                pass

            # 1) Abre o painel (clica no header; se não abrir, tenta o botão de chevron)
            try:
                header.click()
                page.wait_for_timeout(200)
            except Exception:
                pass
            # tenta o botão de expandir se existir
            try:
                chevron_btn = header.locator("xpath=following-sibling::button[1] | xpath=ancestor::*[1]/following-sibling::button[1]").first
                if chevron_btn and chevron_btn.is_visible():
                    chevron_btn.click()
                    page.wait_for_timeout(220)
            except Exception:
                pass

            # 2) Lê tributo e título (como você já fez)
            trib = collapse_spaces(header.locator("p").nth(0).inner_text())
            titulo = collapse_spaces(header.locator("p").nth(1).inner_text())

            # 3) Localiza o bloco de conteúdo do painel com estratégias em cascata
            content = None
            strategies = [
                # irmão seguinte imediato
                "xpath=following-sibling::div[contains(@class,'content')][1]",
                # dentro do wrapper atual (painel aberto costuma ter classe 'open wrapper')
                "xpath=ancestor::*[contains(@class,'wrapper')][1]//div[contains(@class,'content')]",
                # variação: subir um nível e pegar o seguinte .content
                "xpath=../following-sibling::div[contains(@class,'content')][1]",
            ]

            for sel in strategies:
                try:
                    cand = header.locator(sel)
                    if cand and cand.count() > 0:
                        content = cand.first
                        break
                except Exception:
                    continue

            if not content:
                # último recurso: procura o primeiro .content após o header no container das abas
                try:
                    content = page.locator("div.tab-content div.content").nth(i)
                except Exception:
                    content = None

            if not content:
                # Se ainda não localizamos, registra e segue para o próximo painel
                log_warn(f"Alíquota: não encontrei 'content' para o painel {trib}")
                items.append({"trib": trib, "titulo": titulo, "percent": "", "vigencia": "", "base_legal": ""})
                continue

            # 4) Aguarda as linhas ficarem anexadas/visíveis (evita ler antes de renderizar)
            try:
                linhas_first = content.locator(".linha").first
                linhas_first.wait_for(state="visible", timeout=2500)
            except Exception:
                # mesmo que não apareça 'visible' (alguns temas usam overflow), tentamos attached
                try:
                    linhas_first.wait_for(state="attached", timeout=2500)
                except Exception:
                    log_warn(f"Alíquota: '.linha' não visível no painel {trib}")
                    # ainda assim tentaremos extrair abaixo

            # 5) Extrai percentuais, vigências e base legal
            percent = ""
            inicio = ""
            fim = ""
            base_legal_parts: list[str] = []

            try:
                linhas = content.locator(".linha")
                ln = linhas.count()
                for j in range(ln):
                    linha = linhas.nth(j)
                    raw_text = collapse_spaces(linha.inner_text())

                    # Alíquota geral
                    if re.search(r"al[ií]quota geral", raw_text, flags=re.IGNORECASE):
                        m = re.search(r"(\d+[.,]?\d*)\s*%", raw_text)
                        if m and not percent:
                            percent = m.group(1).replace(".", ",") + "%"

                    # Vigência (captura início e fim no mesmo texto)
                    if re.search(r"vig[êe]ncia", raw_text, flags=re.IGNORECASE):
                        # Procura duas datas no mesmo bloco, ex.: 01/01/2026 ... 31/12/2026
                        datas = re.findall(r"(\d{2}/\d{2}/\d{4})", raw_text)
                        if len(datas) >= 2:
                            inicio, fim = datas[0], datas[1]
                        elif len(datas) == 1:
                            # Tenta deduzir se é início ou fim
                            if re.search(r"in[ií]cio", raw_text, flags=re.IGNORECASE) and not inicio:
                                inicio = datas[0]
                            elif re.search(r"fim", raw_text, flags=re.IGNORECASE) and not fim:
                                fim = datas[0]

                    # Base Legal
                    if re.search(r"base legal", raw_text, flags=re.IGNORECASE):
                        links = linha.locator("a")
                        ql = links.count()
                        if ql > 0:
                            for k in range(ql):
                                link_txt = collapse_spaces(links.nth(k).inner_text())
                                if link_txt:
                                    base_legal_parts.append(link_txt)
                        else:
                            # fallback: texto sem link
                            artigos = re.findall(r"Artigo[^;]+", raw_text)
                            for art in artigos:
                                base_legal_parts.append(art.strip())


            except Exception as e:
                log_warn(f"Alíquota: erro ao iterar '.linha' do painel {trib}: {e}")

            vigencia = f"{inicio} - {fim}".strip(" -") if (inicio or fim) else ""
            base_legal = "; ".join([p for p in base_legal_parts if p])

            items.append({
                "trib": trib or "",
                "titulo": titulo or "",
                "percent": percent or "",
                "vigencia": vigencia or "",
                "base_legal": base_legal or "",
            })

    except Exception as e:
        log_err(f"Alíquota: erro inesperado na extração: {e}")

    return items


def extract_pairs_dom(page) -> List[str]:
    """
    Tenta extrair linhas "COD - Descrição" procurando DOM típico:
    <div class="expansion-panel-header"><p>COD</p><p>Descrição</p></div>
    """
    pairs: List[str] = []
    try:
        rows = page.locator("div.tab-content .expansion-panel-header")
        n = rows.count()
        for i in range(n):
            row = rows.nth(i)
            p0 = collapse_spaces(row.locator("p").nth(0).inner_text())
            p1 = collapse_spaces(row.locator("p").nth(1).inner_text())
            if not p0 and not p1:
                continue
            if p0 and p1:
                pairs.append(f"{p0} - {p1}")
            else:
                pairs.append(p0 or p1)
    except Exception as e:
        log_info(f"extract_pairs_dom fallback: {e}")
    return pairs

def extract_pairs_text_fallback(raw_text: str) -> List[str]:
    """
    Se o DOM não ajudar, usa o texto bruto para formar pares:
    remove cabeçalhos e casa padrões como '000' + próxima linha.
    """
    lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
    lines = [l for l in lines if not is_header_line(l)]

    pairs: List[str] = []
    i = 0
    while i < len(lines):
        cur = collapse_spaces(lines[i])
        nxt = collapse_spaces(lines[i+1]) if i+1 < len(lines) else ""
        # códigos mais comuns: 000, 100, 200, 300... ou 2-4 dígitos
        if re.fullmatch(r"\d{2,4}", cur):
            if nxt:
                pairs.append(f"{cur} - {nxt}")
            else:
                pairs.append(cur)
            i += 2
        else:
            # se vier descrição solta, registra como está
            if cur:
                pairs.append(cur)
            i += 1
    return pairs

def extract_structured_items(page, key: str) -> List[str]:
    """
    Para 'CST', 'cClassTrib', 'cCredPres' gera sempre 'COD - Descrição' quando possível.
    Para 'Alíquota', retorna o texto limpo (linhas) sem cabeçalhos.
    """
    # 1) Tenta via DOM
    dom_pairs = extract_pairs_dom(page)
    if dom_pairs:
        log_info(f"{key}: {len(dom_pairs)} item(ns) via DOM.")
        return dom_pairs

    # 2) Fallback por texto
    raw = extract_active_tab_text(page)
    if not raw:
        log_info(f"{key}: conteúdo vazio.")
        return []

    if key in ("CST", "cClassTrib", "cCredPres"):
        pairs = extract_pairs_text_fallback(raw)
        log_info(f"{key}: {len(pairs)} item(ns) via texto.")
        return pairs

    # Alíquota: limpa cabeçalhos e retorna linhas
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    lines = [l for l in lines if not is_header_line(l)]
    log_info(f"{key}: {len(lines)} linha(s) limpas para saída.")
    return lines


def extract_active_tab_text(page) -> str:
    try:
        content = page.locator("div.tab-content").first
        if content and content.is_visible():
            txt = (content.inner_text() or "").strip()
            txt = re.sub(r"\n{3,}", "\n\n", txt)
            return txt
    except Exception:
        pass
    return ""


# ------------------------- WORKFLOW ------------------------
def open_variation_row(page, index: int) -> bool:
    rows = page.locator("div.tab-content table tbody tr")
    count = rows.count()
    if count == 0 or index >= count:
        return False
    row = rows.nth(index)
    try:
        btn = row.locator(".mdi-chevron-right").first
        if btn and btn.is_visible():
            btn.click()
            page.wait_for_timeout(ROW_OPEN_WAIT_MS)
            return True
    except Exception:
        pass
    try:
        row.click()
        page.wait_for_timeout(ROW_OPEN_WAIT_MS)
        return True
    except Exception:
        return False

def ensure_back_to_list(page) -> bool:
    try:
        back_icon = page.locator(".mdi-chevron-left").first
        if back_icon and back_icon.is_visible():
            back_icon.click()
            page.wait_for_timeout(SMALL_WAIT_MS)
    except Exception:
        pass
    if click_tab_by_labels(page, ["NCM", "ncm"]):
        return True
    try:
        page.go_back(wait_until="domcontentloaded")
        if wait_for_variations_table(page, timeout_ms=1200):
            return True
    except Exception:
        pass
    try:
        page.goto(FINAL_LINK, wait_until="domcontentloaded")
        page.wait_for_timeout(SMALL_WAIT_MS)
        click_tab_by_labels(page, ["NCM", "ncm"])
        return True
    except Exception:
        return False

def ensure_ready_for_next_search(page, search_input) -> None:
    try:
        click_tab_by_labels(page, ["NCM", "ncm"])
        page.wait_for_timeout(60)
    except Exception:
        pass
    try:
        search_input.click()
        try:
            search_input.press("Control+A")
        except Exception:
            pass
        try:
            search_input.press("Delete")
        except Exception:
            pass
        search_input.fill("")
        page.wait_for_timeout(60)
    except Exception as e:
        log_warn(f"Falha ao preparar campo de busca para próximo NCM: {e}")


def search_one_ncm(page, search_input, ncm8: str) -> Dict[str, List[str]]:
    t0 = log_step(f"NCM {ncm8}: iniciando busca")
    try:
        search_input.click()
        search_input.fill("")
        time.sleep(0.03)
        search_input.type(ncm8, delay=SEARCH_TYPING_DELAY_MS)
        time.sleep(POST_SEARCH_SLEEP)
        try:
            search_input.press("Enter")
        except Exception:
            pass
        page.wait_for_timeout(SMALL_WAIT_MS)
        log_done(t0, f"NCM {ncm8}: termo enviado")
    except Exception as e:
        log_warn(f"NCM {ncm8}: falha ao digitar/enviar. Erro: {e}")

    t1 = log_step(f"NCM {ncm8}: garantindo aba 'NCM' ativa")
    if click_tab_by_labels(page, ["NCM", "ncm"]):
        log_done(t1, f"NCM {ncm8}: aba 'NCM' ativa")
    else:
        log_warn(f"NCM {ncm8}: não foi possível ativar a aba 'NCM'")

    t2 = log_step(f"NCM {ncm8}: aguardando tabela de variações")
    if not wait_for_variations_table(page, timeout_ms=DEFAULT_WAIT_TABLE_MS):
        log_warn(f"NCM {ncm8}: tabela de variações não apareceu em {DEFAULT_WAIT_TABLE_MS} ms")
        return {"CST": [], "cClassTrib": [], "cCredPres": [], "Alíquota": []}
    log_done(t2, f"NCM {ncm8}: tabela visível")

    rows = page.locator("div.tab-content table tbody tr")
    try:
        total_rows = rows.count()
    except Exception:
        total_rows = 0
    log_info(f"NCM {ncm8}: {total_rows} variação(ões) encontrada(s)")

    aggregated = {"CST": [], "cClassTrib": [], "cCredPres": [], "Alíquota": [], "__ALIQUOTA_STRUCT__": []}

    for i in range(total_rows):
        log_info(f"NCM {ncm8}: abrindo variação {i+1}/{total_rows}")
        click_tab_by_labels(page, ["NCM", "ncm"])
        page.wait_for_timeout(40)

        t_open = log_step(f"NCM {ncm8}: abrir variação {i+1}")
        if not open_variation_row(page, i):
            log_warn(f"NCM {ncm8}: falha ao abrir variação {i+1}")
            continue
        log_done(t_open, f"NCM {ncm8}: variação {i+1} aberta")

        # Aguarda headers das abas
        t_hdr = log_step(f"NCM {ncm8} var {i+1}: aguardando headers de abas")
        if wait_for_header_tabs(page, timeout_ms=DEFAULT_WAIT_HEADERS_MS):
            log_done(t_hdr, f"NCM {ncm8} var {i+1}: headers visíveis")
        else:
            log_warn(f"NCM {ncm8} var {i+1}: headers não apareceram em {DEFAULT_WAIT_HEADERS_MS} ms")

        # Para cada aba alvo, lê badge (com retries) e extrai conteúdo estruturado
        for key, labels in DETAIL_TABS.items():
            badge = read_badge_with_retries(page, labels, retries=DEFAULT_BADGE_RETRIES, sleep_ms=DEFAULT_WAIT_BADGE_RETRY_MS)
            if badge >= 1:
                if click_tab_by_labels(page, labels):
                    page.wait_for_timeout(120)

                    if key == "Alíquota":
                        # >>> NOVO: coleta estruturada por tributo <<<
                        aliqs = extract_aliquota_struct(page)
                        aggregated["__ALIQUOTA_STRUCT__"].extend(aliqs)
                    else:
                        t_click = log_step(f"NCM {ncm8} var {i+1}: clicando aba '{key}'")
                        if click_tab_by_labels(page, labels):
                            log_done(t_click, f"NCM {ncm8} var {i+1}: aba '{key}' ativa")
                            page.wait_for_timeout(120)
                            t_content = log_step(f"NCM {ncm8} var {i+1}: extraindo conteúdo '{key}' (estruturado)")
                            items = extract_structured_items(page, key)
                            if items:
                                log_done(t_content, f"NCM {ncm8} var {i+1}: coletados {len(items)} item(ns) em '{key}'")
                                for it in items:
                                    if it not in aggregated[key]:
                                        aggregated[key].append(it)
                            else:
                                log_warn(f"NCM {ncm8} var {i+1}: nada extraído em '{key}'")
                        else:
                            log_warn(f"NCM {ncm8} var {i+1}: aba '{key}' não clicável")
                else:
                    log_info(f"NCM {ncm8} var {i+1}: aba '{key}' sem itens (badge 0)")

        t_back = log_step(f"NCM {ncm8} var {i+1}: voltando para lista de variações")
        if ensure_back_to_list(page):
            log_done(t_back, f"NCM {ncm8} var {i+1}: voltou para lista")
        else:
            log_warn(f"NCM {ncm8} var {i+1}: falha ao voltar para a lista após várias tentativas")

    ensure_ready_for_next_search(page, search_input)
    return aggregated


def main():
    log("=== ECO CLASS (Econet) - Consulta de NCM via CDP (com Alíquota estruturada) ===")
    log(f"Conectando a: {CDP_ENDPOINT}")
    pw, browser = connect_to_chrome_over_cdp(CDP_ENDPOINT)

    try:
        # 1) Abre/usa a aba do Eco Class
        page = get_or_open_eco_class_page(browser, FINAL_LINK)
        log(f">> Usando a aba: {page.url}")

        # 2) Excel de entrada (salva saída ao lado dele)
        excel_path = ask_excel_path()
        if not excel_path or not os.path.exists(excel_path):
            log_err("Arquivo não encontrado. Encerrando.")
            return
        excel_dir = os.path.dirname(os.path.abspath(excel_path))

        # 3) Lê planilha e detecta colunas
        df = pd.read_excel(excel_path, dtype=str)
        if df.empty:
            log_err("Planilha vazia.")
            return

        codigo_col, nome_col, ncm_col = detect_columns(df)
        if not ncm_col:
            log_err("não localizei NCM no excel selecionado")
            return

        log_info(f"Colunas detectadas -> Código: {codigo_col} | Nome: {nome_col} | NCM: {ncm_col}")

        # 4) Normaliza NCMs (8 dígitos) e deduplica
        df["_NCM8_"] = df[ncm_col].astype(str).map(ncm_to_8digits)

        # Remove NCMs vazios, nulos ou zerados
        unique_ncms = []
        for n in df["_NCM8_"].tolist():
            n_str = str(n).strip()
            if not n_str or n_str in ["0", "00000000", "0000000", "000000", "nan", "None"]:
                continue
            unique_ncms.append(n_str)
        unique_ncms = sorted(set(unique_ncms))

        log_info(f"Total de NCMs únicos: {len(unique_ncms)}")

        # 5) Localiza campo de busca e processa cada NCM
        search_input = find_search_input(page)

        results_by_ncm: Dict[str, Dict[str, list]] = {}
        for idx, ncm8 in enumerate(unique_ncms, 1):
            log_info(f"=== ({idx}/{len(unique_ncms)}) Início NCM {ncm8} ===")
            t_ncm = time.perf_counter()

            if not ncm8 or ncm8 in ["0", "00000000"]:
                log_info(f"NCM vazio/zerado ignorado: {ncm8}")
                continue

            results = search_one_ncm(page, search_input, ncm8)
            results_by_ncm[ncm8] = results
            dt_ncm = (time.perf_counter() - t_ncm) * 1000.0
            log_info(f"=== Fim NCM {ncm8} (em {dt_ncm:.0f} ms) ===")

            # --- LIMPEZA PERIÓDICA DE COOKIES DA ECONET ---
            if idx % 20 == 0:
                log_info("Atingidos 20 NCMs — limpando cookies da Econet (preservando sessão).")
                clear_econet_cookies_keep_session(page)

        # 6) Descobre todos os TRIBUTOS encontrados na aba Alíquota (IBS, CBS, ...)
        all_tribs: list[str] = []
        for _ncm8, res in results_by_ncm.items():
            aliqs = res.get("__ALIQUOTA_STRUCT__", [])
            for it in aliqs:
                trib = (it.get("trib", "") or "").strip()
                if trib and trib not in all_tribs:
                    all_tribs.append(trib)

        # 7) Monta DataFrame base
        final_df = pd.DataFrame()
        final_df["Codigo do produto"] = df[codigo_col].astype(str) if (codigo_col and codigo_col in df.columns) else ""
        final_df["Nome"] = df[nome_col].astype(str) if (nome_col and nome_col in df.columns) else ""
        final_df["NCM"] = df[ncm_col].astype(str)

        # 8) Preenche abas simples (sem mudar sua lógica atual)
        def join_or_blank(ncm8: str, key: str) -> str:
            pack = results_by_ncm.get(ncm8, {}).get(key, [])
            return "\n".join(pack) if pack else ""

        final_df["CST"] = df["_NCM8_"].map(lambda n: join_or_blank(n, "CST"))
        final_df["cClassTrib"] = df["_NCM8_"].map(lambda n: join_or_blank(n, "cClassTrib"))
        final_df["cCredPres"] = df["_NCM8_"].map(lambda n: join_or_blank(n, "cCredPres"))

        # 9) Cria colunas dinâmicas para cada tributo da aba Alíquota
        for trib in all_tribs:
            final_df[f"Alíquota {trib}"] = ""
            final_df[f"% {trib}"] = ""
            final_df[f"Vigencia {trib}"] = ""
            final_df[f"Base Legal {trib}"] = ""

        # 10) Preenche, linha a linha, as colunas de Alíquota com o que foi coletado em "__ALIQUOTA_STRUCT__"
        records = []
        for idx_row, row in final_df.iterrows():
            ncm8 = df.iloc[idx_row]["_NCM8_"]
            rec = row.to_dict()

            aliqs = results_by_ncm.get(ncm8, {}).get("__ALIQUOTA_STRUCT__", [])
            for it in aliqs:
                trib = (it.get("trib", "") or "").strip()
                if not trib:
                    continue
                rec[f"Alíquota {trib}"] = it.get("titulo", "")
                rec[f"% {trib}"] = it.get("percent", "")
                rec[f"Vigencia {trib}"] = it.get("vigencia", "")
                rec[f"Base Legal {trib}"] = it.get("base_legal", "")

            records.append(rec)

        final_df = pd.DataFrame(records)

        # 11) (Opcional) Ordena colunas deixando as novas ao final em grupos por tributo
        ordered = ["Codigo do produto", "Nome", "NCM", "CST", "cClassTrib", "cCredPres"]
        for trib in all_tribs:
            ordered += [f"Alíquota {trib}", f"% {trib}", f"Vigencia {trib}", f"Base Legal {trib}"]
        for c in final_df.columns:
            if c not in ordered:
                ordered.append(c)
        final_df = final_df[ordered]

        # 12) Salva ao lado do Excel original
        # Nome base do excel de entrada (sem caminho e sem extensão)
        original_base = os.path.splitext(os.path.basename(excel_path))[0]

        # Novo nome: Resultado "nome do excel original" CCLASS.xlsx
        out_name = f'Resultado {original_base} CCLASS.xlsx'

        out_path = os.path.join(excel_dir, out_name)
        final_df.to_excel(out_path, index=False, engine="openpyxl")
        log(f"Concluído! Arquivo gerado ao lado do Excel original: {out_path}")

    finally:
        try:
            pw.stop()
        except Exception:
            pass



if __name__ == "__main__":
    main()
