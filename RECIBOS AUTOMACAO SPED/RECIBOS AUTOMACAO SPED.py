import asyncio
import os
import time
import re
import traceback
import unicodedata
from html import unescape
from datetime import datetime, timedelta
from pathlib import Path
from re import compile as re_compile, IGNORECASE

from CHROME_9222_CAPTCHA import get_browser, PORT
from playwright.async_api import async_playwright, TimeoutError as PWTimeout, Error as PWError

# ---------------------------
# Configurações / Credenciais
# ---------------------------
BASE_DIR = Path(__file__).resolve().parent


def load_local_env(*paths: Path) -> None:
    for path in paths:
        if not path.is_file():
            continue
        try:
            with path.open("r", encoding="utf-8") as handle:
                for line in handle:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()
                    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
                        value = value[1:-1]
                    if key and key not in os.environ:
                        os.environ[key] = value
        except OSError:
            continue


load_local_env(BASE_DIR / ".env")

USERNAME = os.getenv("IOB_USUARIO", "").strip()
PASSWORD = os.getenv("IOB_SENHA", "").strip()

DOWNLOAD_DIR = r"C:\Users\Usuario\Downloads\RECIBOS_SPED"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

LOG_FILE = os.path.join(DOWNLOAD_DIR, "log.txt")

LOGIN_URL = os.getenv(
    "IOB_LOGIN_URL",
    (
        "https://sso.iob.com.br/signin/?response_type=code&scope=&client_id=b89f6d4c-78bb-4995-9ac7-aa5459f5cf6d"
        "&redirect_uri=https%3A%2F%2Fapp.iob.com.br%2Fcallback%2F%3Fpath%3Dhttps%3A%2F%2Fapp.iob.com.br%2Fapp%2F"
        "&isSignUpDisable=false&showFAQ=false&isSocialLoginDisable=false&title=Com+um+s%C3%B3+acesso%2C+voc%C3%AA+se+conecta+%C3%A0+tecnologia+e+intelig%C3%AAncia+IOB"
        "&subtitle=Que+bom+ter+voc%C3%AA+aqui%21"
    ),
)

CAPTCHA_IFRAME_SEL = "iframe[title*='reCAPTCHA'], iframe[src*='recaptcha'], iframe[src*='hcaptcha']"
ACCESS_BTN_TEXT = "Acessar aqui"
APP_URL_RE = re_compile(r"^https://app\.iob\.com\.br/.*", IGNORECASE)

STATUS_BADGE_SEL = "div.css-feuuq5, div.css-qwusia, :text('Transmissão Concluída'), :text('Cancelado')"
COMP_LABEL_SEL = "span:has-text('Competência:')"

DOCS_BASE = "https://docs.iob.com.br/api"


# =============== LOGGING ===============
_LOG_CTX = {"item": "", "empresa": ""}


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]


def set_ctx(item: str = "", empresa: str = ""):
    _LOG_CTX["item"] = item
    _LOG_CTX["empresa"] = empresa


def log(msg: str):
    line = f"[{_ts()}][{_LOG_CTX.get('item','')}]({_LOG_CTX.get('empresa','')}) {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def log_exc(prefix: str = "EXCEPTION"):
    tb = traceback.format_exc()
    log(f"{prefix}: {tb.strip()}")


# =============== Utils de nomes ===============

def strip_accents(s: str) -> str:
    return unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode("ascii")


def clean_filename(name: str) -> str:
    invalid = r'<>:"/\\|?*\n\r\t'
    for ch in invalid:
        name = name.replace(ch, "_")
    return name.strip()


def clean_filename_strict(name: str) -> str:
    name = strip_accents(name).strip()
    for ch in r'<>:"/\|?*':
        name = name.replace(ch, "_")
    name = re.sub(r"\s+", " ", name).strip()
    return name


async def build_pdf_target_path(download_dir: str, card_title_text: str) -> str:
    base = f"{clean_filename_strict(card_title_text)} - EFD Contribuicoes.pdf"
    return os.path.join(download_dir, base)


# =============== Login / Navegação ===============

async def wait_captcha_with_retries(page, attempts=3, per_try_timeout_ms=10000):
    for i in range(1, attempts + 1):
        try:
            await page.wait_for_selector(CAPTCHA_IFRAME_SEL, timeout=per_try_timeout_ms)
            log(f"Captcha detectado (tentativa {i}/{attempts}).")
            return True
        except PWTimeout:
            log(f"Captcha não apareceu (tentativa {i}/{attempts}). Recarregando a página…")
            try:
                await page.reload(wait_until="load")
            except Exception as e:
                log(f"Aviso ao recarregar: {e}")
    return False


def app_url_ok(u: str) -> bool:
    return ("app.iob.com.br" in u) or ("callback" in u)


async def after_submit_pick_page(context, current_page, url_regex=APP_URL_RE, timeout=60000):
    log("after_submit_pick_page: aguardando aba correta…")

    async def wait_same():
        await current_page.wait_for_url(url_regex, timeout=timeout)
        try:
            await current_page.wait_for_load_state("load", timeout=15000)
        except Exception:
            pass
        return current_page

    async def wait_new():
        new_page = await context.wait_for_event("page", timeout=timeout)
        try:
            await new_page.wait_for_url(url_regex, timeout=30000)
        except Exception:
            pass
        try:
            await new_page.wait_for_load_state("load", timeout=30000)
        except Exception:
            pass
        if not current_page.is_closed():
            try:
                await current_page.close()
            except Exception:
                pass
        return new_page

    async def wait_closed():
        await current_page.wait_for_event("close", timeout=timeout)
        if not context.pages:
            raise PWError("Página fechada após login e nenhuma outra aba disponível.")
        page = context.pages[-1]
        try:
            await page.wait_for_load_state("load", timeout=30000)
        except Exception:
            pass
        return page

    tasks = [
        asyncio.create_task(wait_same()),
        asyncio.create_task(wait_new()),
        asyncio.create_task(wait_closed()),
    ]
    done, pending = await asyncio.wait(tasks, return_when=asyncio.FIRST_COMPLETED)
    winner = next(iter(done))
    for t in pending:
        t.cancel()
    await asyncio.gather(*pending, return_exceptions=True)
    log("after_submit_pick_page: aba detectada.")
    return winner.result()


async def goto_painel_sped(page):
    if "app.iob.com.br/app/sped/painel" not in page.url:
        log("Navegando para /app/sped/painel…")
        await page.goto("https://app.iob.com.br/app/sped/painel/", wait_until="networkidle")
    await page.wait_for_load_state("domcontentloaded")
    await page.wait_for_timeout(900)
    log("Painel SPED carregado (domcontentloaded).")


async def get_painel_concluido(page, timeout_ms=60000):
    log("Procurando painel 'Concluído'…")
    titulo = await page.wait_for_selector(
        "h6.MuiTypography-subtitle1:has-text('Concluído')",
        timeout=timeout_ms
    )
    h = await titulo.evaluate_handle(
        "el => el.closest('div.MuiPaper-root.MuiPaper-outlined.MuiPaper-rounded')"
    )
    painel_el = h.as_element()
    if painel_el:
        log("Painel 'Concluído' encontrado via closest.")
        return painel_el

    for cand in await page.query_selector_all(
        "div.MuiPaper-root.MuiPaper-outlined.MuiPaper-rounded"
    ):
        try:
            if await cand.query_selector("h6.MuiTypography-subtitle1:has-text('Concluído')"):
                log("Painel 'Concluído' encontrado via varredura.")
                return cand
        except:
            pass

    possivel = await page.query_selector(":text('Concluído')")
    if possivel:
        h2 = await possivel.evaluate_handle("el => el.closest('div.MuiPaper-root')")
        painel_el2 = h2.as_element()
        if painel_el2:
            log("Painel 'Concluído' encontrado via fallback.")
            return painel_el2

    raise TimeoutError("Painel 'Concluído' não localizado por nenhum método.")


# =============== Cards / filtros ===============

def competencia_mes_anterior() -> str:
    today = datetime.today()
    first = today.replace(day=1)
    prev = first - timedelta(days=1)
    return f"{prev.month:02d}/{prev.year}"


async def card_status(card):
    try:
        lbl = await card.query_selector(STATUS_BADGE_SEL)
        if lbl:
            t = (await lbl.inner_text()).strip()
            if "Concluída" in t or "Concluida" in t:
                return "Transmissão Concluída"
            if "Cancelado" in t:
                return "Cancelado"
    except:
        pass
    try:
        t = (await card.inner_text()).lower()
        if "transmissão concluída" in t or "transmissao concluida" in t:
            return "Transmissão Concluída"
        if "cancelado" in t:
            return "Cancelado"
    except:
        pass
    return None


async def card_competencia(card):
    try:
        label = await card.query_selector(COMP_LABEL_SEL)
        if label:
            val = await label.evaluate_handle(
                "el => (el.parentElement && el.parentElement.querySelectorAll('span')[1]) || null"
            )
            if val:
                txt = await (await val.as_element()).inner_text()
                return txt.strip()
    except:
        pass
    try:
        txt = await card.inner_text()
        m = re.search(r"\b(\d{2}/\d{4})\b", txt or "")
        if m:
            return m.group(1)
    except:
        pass
    return None


# =============== Modal helpers ===============

async def robust_close_modal(page, timeout_ms=4000):
    """
    Fecha o modal por botões conhecidos, depois tenta ESC e click no backdrop.
    (Tudo DOM / Playwright, nada de mouse/teclado do sistema).
    """
    try:
        modal = page.locator(".MuiDialog-root").last
        if await modal.count() == 0:
            return

        # 1) botões padrão
        for sel in ["button:has-text('Fechar')", "[aria-label='Fechar']", "[data-testid='CloseIcon']"]:
            btn = modal.locator(sel).first
            if await btn.count():
                try:
                    await btn.click(force=True, timeout=1200)
                    break
                except Exception:
                    pass

        # 2) ESC (evento DOM)
        try:
            await page.keyboard.press("Escape")
        except Exception:
            pass

        # 3) click no backdrop via DOM
        try:
            back = page.locator(".MuiBackdrop-root").last
            if await back.count():
                try:
                    await back.click(timeout=1000)
                except Exception:
                    pass
        except Exception:
            pass

        # aguardar sumir
        try:
            await page.locator(".MuiDialog-root").wait_for(state="hidden", timeout=timeout_ms)
        except Exception:
            try:
                await page.locator(".MuiDialog-root").wait_for(state="detached", timeout=1500)
            except Exception:
                pass
    except Exception:
        pass


async def open_card_and_wait_modal(page, card, max_tries=3, open_timeout_ms=15000):
    """
    Clica no card e bloqueia até o diálogo abrir.
    """
    click_targets = [
        await card.query_selector("h6.MuiTypography-subtitle2"),
        await card.query_selector(".MuiCardContent-root"),
        card
    ]
    for attempt in range(1, max_tries + 1):
        for tgt in [t for t in click_targets if t]:
            try:
                await tgt.scroll_into_view_if_needed()
            except Exception:
                pass
            try:
                await tgt.click(force=True, timeout=1200)
                log(f"Clique no card (tentativa {attempt}) efetuado.")
                break
            except Exception:
                continue

        try:
            modal_root = await page.wait_for_selector(".MuiDialog-root", timeout=open_timeout_ms)
            if modal_root:
                paper = await modal_root.query_selector(".MuiDialog-paper,.MuiPaper-root")
                return paper or modal_root
        except Exception:
            pass

    return None


# =============== Workflow + Documents (API) ===============

def extract_recibo_document_from_workflow(wf_json: dict) -> dict | None:
    """
    A partir do JSON de /api/workflows/processes/{id}?fullDetails=true,
    encontra o Document que corresponde ao recibo (Recibo de Entrega / comprovante).
    Usa uma heurística de "score" para pegar o melhor candidato.
    """
    if not wf_json:
        return None

    states = wf_json.get("States") or []
    best_doc = None
    best_score = 0

    for st in states:
        attachments = st.get("Attachments") or []
        for att in attachments:
            doc = (att or {}).get("Document") or {}
            name = doc.get("name") or ""
            ext = (doc.get("extension") or "").lower()
            mime = (doc.get("mimeType") or "").lower()
            is_pdf = doc.get("isPdf") or ext == "pdf" or "application/pdf" in mime
            if not is_pdf:
                continue

            norm = strip_accents(name).lower()
            score = 0

            # Prioridade máxima: "recibo" e "entrega" juntos
            if "recibo" in norm and "entrega" in norm:
                score = 100
            # Só "recibo"
            elif "recibo" in norm:
                score = 80

            # "comprovante" (com ou sem entrega)
            if "comprovante" in norm:
                score = max(score, 70)
                if "entrega" in norm or "transmissao" in norm:
                    score = max(score, 90)

            # EFD_CONTRIBUICOES soma um tiquinho de prioridade
            if "efd_contribuicoes" in norm:
                score += 5

            if score > best_score:
                best_score = score
                best_doc = doc

    if not best_doc:
        # loga para debug: quais PDFs existiam?
        try:
            debug_names = []
            for st in states:
                for att in (st.get("Attachments") or []):
                    doc = (att or {}).get("Document") or {}
                    if not doc:
                        continue
                    n = doc.get("name") or ""
                    ext = (doc.get("extension") or "").lower()
                    mime = (doc.get("mimeType") or "").lower()
                    is_pdf = doc.get("isPdf") or ext == "pdf" or "application/pdf" in mime
                    if is_pdf:
                        debug_names.append(n)
            log(f"extract_recibo_document_from_workflow: nenhum recibo claro encontrado. PDFs vistos: {debug_names}")
        except Exception:
            pass

    return best_doc


async def download_pdf_document(
    api_request,
    document_id: int | str,
    card_title: str,
    download_dir: str,
    max_retries: int = 2
) -> str | None:
    """
    Faz GET em https://docs.iob.com.br/api/documents/{id}/download
    usando o browser_context.request (compartilha cookies da sessão),
    e salva como "<TÍTULO CARD> - EFD Contribuicoes.pdf".

    Só aceita a resposta se o conteúdo for realmente PDF
    (content-type OU magic bytes "%PDF-"). Senão tenta novamente
    e por fim salva um .html de debug, marcando como falha.
    """
    try:
        doc_id = str(document_id)
    except Exception:
        log(f"download_pdf_document: document_id inválido: {document_id!r}")
        return None

    url = f"{DOCS_BASE}/documents/{doc_id}/download"
    base_name = f"{clean_filename_strict(card_title or 'arquivo')} - EFD Contribuicoes.pdf"
    full_pdf_path = os.path.join(download_dir, base_name)

    for attempt in range(1, max_retries + 1):
        log(f"download_pdf_document: GET {url} (tentativa {attempt}/{max_retries})")

        try:
            resp = await api_request.get(url)
        except Exception as e:
            log(f"download_pdf_document: erro de rede: {e}")
            return None

        ct = (resp.headers.get("content-type") or "").lower()
        log(f"download_pdf_document: status={resp.status}, content-type={ct}")

        # Se status não OK, loga um trechinho do corpo e desiste ou tenta de novo
        if not resp.ok:
            try:
                body_snip = (await resp.text())[:300]
            except Exception:
                body_snip = "<não foi possível ler corpo>"
            log(f"download_pdf_document: HTTP {resp.status}. Corpo (início): {body_snip}")
            if attempt < max_retries:
                await asyncio.sleep(1.0)
                continue
            return None

        try:
            content = await resp.body()
        except Exception as e:
            log(f"download_pdf_document: erro ao ler body: {e}")
            if attempt < max_retries:
                await asyncio.sleep(1.0)
                continue
            return None

        size = len(content or b"")
        log(f"download_pdf_document: recebidos {size} bytes.")

        # Heurística forte de PDF
        magic_pdf = content.startswith(b"%PDF-")
        is_pdf_ct = "application/pdf" in ct or "application/octet-stream" in ct

        if not (magic_pdf or is_pdf_ct):
            # Muito provavelmente é HTML (viewer) ou erro
            preview = (content[:200] or b"").decode("latin1", errors="ignore")
            log(f"download_pdf_document: conteúdo não parece PDF. Magic={magic_pdf}, content-type={ct}, preview={preview[:120]!r}")
            if attempt < max_retries:
                await asyncio.sleep(1.0)
                continue

            # Última tentativa falhou → salva como .html de debug (sem sobrescrever o .pdf)
            debug_name = f"{clean_filename_strict(card_title or 'arquivo')}_NAO_E_PDF_debug.html"
            debug_path = os.path.join(download_dir, debug_name)
            try:
                with open(debug_path, "wb") as f:
                    f.write(content)
                log(f"download_pdf_document: conteúdo salvo como debug HTML: {debug_path}")
            except Exception as e:
                log(f"download_pdf_document: falha ao salvar HTML de debug: {e}")
            return None

        # OK, parece PDF mesmo
        try:
            with open(full_pdf_path, "wb") as f:
                f.write(content)
            log(f"download_pdf_document: PDF salvo em {full_pdf_path}")
            return full_pdf_path
        except Exception as e:
            log(f"download_pdf_document: erro ao salvar arquivo: {e}")
            if attempt < max_retries:
                await asyncio.sleep(1.0)
                continue
            return None

    return None


async def open_card_and_capture_workflow_details(page, card, timeout_ms: int = 15000):
    """
    Abre o card e, no mesmo movimento, captura a response de
      https://docs.iob.com.br/api/workflows/processes/{id}?fullDetails=true

    Retorna (modal_root, workflow_json) ou (None, None) se não conseguir.
    """

    def is_full_details_response(res):
        try:
            url = res.url or ""
        except Exception:
            return False
        return (
            res.request.method == "GET"
            and "docs.iob.com.br/api/workflows/processes/" in url
            and "fullDetails=true" in url
        )

    try:
        async with page.expect_response(is_full_details_response, timeout=timeout_ms) as resp_info:
            modal_root = await open_card_and_wait_modal(page, card, max_tries=2, open_timeout_ms=timeout_ms)
        resp = await resp_info.value
        try:
            wf_json = await resp.json()
        except Exception as e:
            log(f"Falha ao converter workflow response em JSON: {e}")
            wf_json = None

        return modal_root, wf_json
    except Exception as e:
        log(f"open_card_and_capture_workflow_details: falha ao capturar workflow: {e}")
        return None, None


# ---------------------------
# Main
# ---------------------------

async def run():
    # inicia log de sessão
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 80 + f"\nSessão: {_ts()}\n" + "=" * 80 + "\n")
    except Exception:
        pass

    async with async_playwright() as p:
        browser = await get_browser(p, PORT)

        if not browser.contexts:
            raise RuntimeError("Nenhum contexto persistente disponível no Chrome conectado.")

        context = browser.contexts[0]
        page = context.pages[0] if context.pages else await context.new_page()

        # 1) LOGIN
        log("Abrindo URL de login…")
        await page.goto(LOGIN_URL, wait_until="load")

        captcha_ok = await wait_captcha_with_retries(page, attempts=3, per_try_timeout_ms=10000)
        if captcha_ok:
            try:
                if await page.query_selector("#username"):
                    await page.fill("#username", USERNAME)
                    log("Campo usuário preenchido.")
                if await page.query_selector("#password"):
                    await page.fill("#password", PASSWORD)
                    log("Campo senha preenchido.")
                log("Campos preenchidos automaticamente.")
            except Exception as e:
                log(f"Falha ao preencher campos: {e}")

        await page.wait_for_timeout(400)

        # foco no recaptcha com TAB e SPACE (DOM, não sistema)
        for _ in range(3):
            try:
                await page.keyboard.press("Tab")
            except Exception:
                pass
            await page.wait_for_timeout(300)
        try:
            await page.keyboard.press("Space")
        except Exception:
            pass

        log("Aguardando reCAPTCHA ser resolvido…")
        await page.wait_for_function(
            "(()=>{const ta=document.querySelector('#g-recaptcha-response');"
            "return !!(ta && ta.value && ta.value.trim().length>0);})()",
            timeout=180_000
        )
        log("reCAPTCHA resolvido; clicando em Entrar…")

        # clicar Entrar
        try:
            sel_btn_login = "button#formButton, button[name='login'], button.primary-subscription:has-text('Entrar')"
            await page.wait_for_selector(sel_btn_login, timeout=30_000)
            login_btn = page.locator(sel_btn_login).first
            try:
                await login_btn.click(timeout=8_000, trial=True)
                await login_btn.click(timeout=8_000)
            except Exception:
                try:
                    await login_btn.scroll_into_view_if_needed()
                    await login_btn.click(timeout=8_000, force=True)
                except Exception:
                    try:
                        await page.keyboard.press("Enter")
                    except Exception:
                        pass
            page = await after_submit_pick_page(context, page, url_regex=APP_URL_RE, timeout=45000)
            log("Login enviado / página de app detectada.")
        except Exception:
            log_exc("Falha ao acionar login")

        # 2) /app
        try:
            if not app_url_ok(page.url):
                log("Forçando /app …")
                await page.goto("https://app.iob.com.br/app/", wait_until="load")
            await page.wait_for_load_state("domcontentloaded", timeout=20000)
            await page.wait_for_timeout(800)
        except Exception:
            pass

        time.sleep(2.0)

        # 3) "Acessar aqui"
        try:
            sel_access = f"button:has-text('{ACCESS_BTN_TEXT}'), .MuiButton-root:has-text('{ACCESS_BTN_TEXT}')"
            btn = page.locator(sel_access).first
            await btn.wait_for(state="visible", timeout=10_000)
            handle = await btn.element_handle()
            await page.wait_for_function(
                "(el)=>!!el && !el.disabled && !el.getAttribute('aria-disabled')",
                arg=handle, timeout=6_000
            )
            await btn.click(timeout=8_000)
            try:
                await page.wait_for_load_state("load", timeout=20_000)
            except Exception:
                pass
            log("'Acessar aqui' clicado.")
        except Exception:
            log("'Acessar aqui' não apareceu ou falhou.")

        time.sleep(2.0)

        # 4) /sped/painel
        try:
            log("Indo para /sped/painel…")
            await page.goto("https://app.iob.com.br/app/sped/painel/", wait_until="load")
        except Exception:
            pass

        # 5) painel "Concluído"
        try:
            await goto_painel_sped(page)
            concluido_panel = await get_painel_concluido(page, timeout_ms=60000)

            CARD_SEL_LOCAL = "div.MuiPaper-root.MuiPaper-elevation.MuiPaper-rounded.MuiPaper-elevation0"
            # scroll até o fim para carregar todos os cards
            for _ in range(14):
                try:
                    await concluido_panel.evaluate("(el) => { el.scrollTop = el.scrollHeight }")
                except Exception:
                    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                await page.wait_for_timeout(160)

            await concluido_panel.wait_for_selector(CARD_SEL_LOCAL, timeout=60000)
            cards = await concluido_panel.query_selector_all(CARD_SEL_LOCAL)
            if not cards:
                log("Nenhum card encontrado no painel 'Concluído'.")
                return
            log(f"Cards no painel 'Concluído': {len(cards)}")
        except Exception as e:
            log(f"Aviso: não foi possível localizar o painel 'Concluído' no tempo esperado. ({e})")
            return

        # timeouts mais curtos durante o loop
        page.set_default_timeout(5000)

        alvo_competencia = competencia_mes_anterior()
        log(f"Filtrando por competência do mês anterior: {alvo_competencia}")

        # Loop dos cards
        for i, card in enumerate(cards, start=1):
            set_ctx(item=f"card#{i}", empresa="")
            try:
                # nome da empresa
                try:
                    name_el = await card.query_selector("h6.MuiTypography-subtitle2, .MuiTypography-subtitle2")
                    empresa = (await name_el.inner_text()).strip() if name_el else f"Item{i}"
                    empresa = unescape(empresa).strip()
                except Exception:
                    empresa = f"Item{i}"
                set_ctx(item=f"card#{i}", empresa=empresa)

                st = await card_status(card)
                comp = await card_competencia(card) or "N/A"
                log(f"Status={st or 'N/A'} | Competência={comp}")

                if st != "Transmissão Concluída" or comp != alvo_competencia:
                    log("Card pulado (não válido).")
                    continue

                # título do card para o nome do arquivo
                try:
                    title_el = await card.query_selector("h6.MuiTypography-subtitle2, .MuiTypography-subtitle2")
                    card_title = (await title_el.inner_text()).strip() if title_el else empresa
                except Exception:
                    card_title = empresa

                target_path_preview = await build_pdf_target_path(DOWNLOAD_DIR, card_title)
                log(f"Nome-alvo para o PDF: {target_path_preview}")

                log("Card válido → abrir overlay e capturar workflow details…")
                modal_root, wf_json = await open_card_and_capture_workflow_details(
                    page, card, timeout_ms=15000
                )

                if not modal_root:
                    log("NÃO abriu dialog ou não capturou workflow; seguindo para o próximo card.")
                    continue

                if not wf_json:
                    log("Workflow JSON vazio/nulo; tentando fechar modal e seguir.")
                    await robust_close_modal(page)
                    await page.wait_for_timeout(200)
                    continue

                # extrai o Document do Recibo
                doc = extract_recibo_document_from_workflow(wf_json)
                if not doc:
                    log("Não foi possível encontrar o Document de recibo no workflow.")
                    await robust_close_modal(page)
                    await page.wait_for_timeout(200)
                    continue

                doc_id = doc.get("id")
                doc_name = doc.get("name")
                log(f"Document de recibo encontrado: id={doc_id}, name={doc_name}")

                # fecha o modal antes de baixar
                await robust_close_modal(page)
                await page.wait_for_timeout(200)

                # baixa o PDF direto pela API documents/{id}/download
                saved_path = await download_pdf_document(
                    context.request,
                    document_id=doc_id,
                    card_title=card_title,
                    download_dir=DOWNLOAD_DIR
                )

                if saved_path:
                    log(f"PDF salvo com sucesso: {saved_path}")
                else:
                    log("Falha ao salvar PDF via API para este card.")

            except Exception:
                log_exc("Erro no loop de card")
                try:
                    await robust_close_modal(page)
                except Exception:
                    pass
                continue

        log("Processamento finalizado.")
        print(f"Log salvo em: {LOG_FILE}")


if __name__ == "__main__":
    asyncio.run(run())
