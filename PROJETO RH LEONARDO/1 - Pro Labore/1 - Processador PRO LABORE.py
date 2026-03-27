import argparse
import ctypes
import ctypes.wintypes as wintypes
import json
import os
import re
import unicodedata
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path

try:
    from pywinauto import Desktop, mouse
    from pywinauto.keyboard import send_keys as send_keys_raw
except Exception as exc:
    raise SystemExit("pywinauto is required for UI automation") from exc

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None
try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None


# =========================
# CONFIG
# =========================

ATTACH_TITLE_CONTAINS = "folha de pagamento"
PROCESSAMENTO_VALUE = "2"
MES_ANO_VALUE = datetime.now().strftime("%m/%Y")
GERAR_POR_VALUE = "1"
GRUPO_VALUE = "1"

RAW_DIR_PRIMARY = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\1 - Pro Labore\Arquivos Brutos"
RAW_DIR_FALLBACK = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\1 - Pro Labore\Arquivos Brutos"

OUTPUT_ROOT = r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\PRO LABORE"
IDLE_SECONDS = 30

COORDS_PATH = os.path.join(os.path.dirname(__file__), "coords_prolabore.json")

KEY_PAUSE = 0.04
CLICK_DELAY = 0.2
FIELD_DELAY = 0.2
POLL_SECONDS = 5
PRINT_OFFER_POLL = 1.5
FINAL_OK_COORD = (1080, 600)
FINAL_CLOSE_COORD = (1471, 169)
FINAL_CLOSE_WAIT_S = 2.0
FINAL_BETWEEN_WAIT_S = 15.0
UI_CLICK_GAP_S = 1.0


COORDS = {
    "field_processamento": None,
    "field_mes_ano": None,
    "field_gerar_por": None,
    "field_grupo": None,
    "ctx_calc_folha": None,
    "ctx_apuracao": None,
    "ctx_holerite": None,
    "ctx_extra": None,
    "menu_parametros": None,
    "menu_relatorio_mensal": None,
    "btn_diretorio": None,
    "btn_selecionar_pasta": None,
    "btn_processar": None,
}


# =========================
# UTIL
# =========================

def log(msg: str):
    print(msg, flush=True)


def load_coords(path: str) -> dict:
    if not path or not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        return {}
    return data


def save_coords(path: str, data: dict):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def send_keys(text: str, pause: float = KEY_PAUSE):
    send_keys_raw(text, pause=pause, with_spaces=True)


def click_abs(pos, pre_wait_s: float = UI_CLICK_GAP_S):
    mouse.move(coords=pos)
    time.sleep(pre_wait_s)
    mouse.click(button="left", coords=pos)
    time.sleep(CLICK_DELAY)


def replace_text_at(coord, text: str):
    click_abs(coord)
    time.sleep(FIELD_DELAY)
    send_keys("^a")
    send_keys(text)
    time.sleep(FIELD_DELAY)


def context_menu_sequence(coord, moves: str, delay_s: float = 0.2, post_delay_s: float = 0.5):
    mouse.click(button="right", coords=coord)
    time.sleep(delay_s)
    for key in moves:
        if key == "D":
            send_keys("{DOWN}")
        elif key == "R":
            send_keys("{RIGHT}")
        elif key == "E":
            send_keys("{ENTER}")
        time.sleep(delay_s)
    time.sleep(post_delay_s)


def get_raw_dir() -> str:
    if os.path.isdir(RAW_DIR_PRIMARY):
        return RAW_DIR_PRIMARY
    return RAW_DIR_FALLBACK


def wait_file_stable(path: Path, timeout_s: float = 45.0, stable_for_s: float = 2.0) -> bool:
    t0 = time.time()
    last_size = None
    stable_since = None
    while (time.time() - t0) < timeout_s:
        if not path.exists():
            time.sleep(0.2)
            continue
        size = path.stat().st_size
        if last_size is None or size != last_size:
            last_size = size
            stable_since = time.time()
        else:
            if stable_since and (time.time() - stable_since) >= stable_for_s:
                return True
        time.sleep(0.4)
    return False


def sanitize_filename(name: str) -> str:
    invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid:
        name = name.replace(ch, "-")
    name = " ".join(name.split())
    return name.strip()


def _extract_company_from_line(line: str) -> str:
    text = (line or "").strip()
    if not text:
        return ""
    lowered = normalize_pdf_text(text)
    for token in ("empresa:", "razao social:"):
        if token in lowered:
            parts = text.split(":", 1)
            if len(parts) == 2:
                return parts[1].strip()
    return text


def normalize_pdf_text(text: str) -> str:
    raw = (text or "").strip().lower()
    return "".join(ch for ch in unicodedata.normalize("NFKD", raw) if not unicodedata.combining(ch))


def _is_blacklisted_line(line: str) -> bool:
    lowered = normalize_pdf_text(line)
    return any(phrase in lowered for phrase in BLACKLIST_PHRASES)


BLACKLIST_PHRASES = [
    "recibo de retiradas",
    "recibo de retirada",
    "recibo",
    "pro labore",
    "holerith",
    "holerite",
    "folha",
    "competencia",
    "competência",
    "referente",
    "cpf",
    "cnpj",
    "matricula",
    "cod.",
    "cód.",
    "codigo",
    "código",
    "total de vencimentos",
    "total vencimentos",
    "relatorio",
    "relatório",
    "processamento",
    "trial mode",
    "click here for more information",
    "codigo nome do funcionario",
    "cbo",
    "local depto",
    "setor",
    "secao",
    "seção",
    "descricao",
    "descrição",
    "sal. contr. inss",
    "nome",
    "emp.",
]

COMPANY_TOKENS = [
    "ltda",
    "eireli",
    "s/a",
    "sa",
    "me",
    "epp",
    "mei",
    "holding",
    "industria",
    "comercio",
    "servicos",
    "sociedade",
    "advocacia",
    "advogados",
    "clinica",
    "odontologia",
    "fisioterapia",
    "participacoes",
    "administracao",
    "contabilidade",
    "corretora",
    "assessoria",
    "representacoes",
    "equipamentos",
    "construcao",
]

CNPJ_RE = re.compile(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b")
CPF_RE = re.compile(r"\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b")


def _has_company_token(line: str) -> bool:
    lowered = normalize_pdf_text(line)
    return any(tok in lowered for tok in COMPANY_TOKENS)


def _is_employee_marker(line: str) -> bool:
    lowered = normalize_pdf_text(line)
    return any(word in lowered for word in ("funcionario", "empregado", "colaborador", "nome"))


def _looks_like_person(line: str) -> bool:
    text = (line or "").strip()
    if not text:
        return False
    if any(ch.isdigit() for ch in text):
        return False
    if _has_company_token(text) or _is_employee_marker(text):
        return False
    words = [w for w in text.replace(".", " ").split() if w]
    if len(words) < 2 or len(words) > 4:
        return False
    cap_words = sum(1 for w in words if w[:1].isupper())
    return cap_words >= len(words) - 1


def _read_first_page_text(pdf_path: Path) -> str:
    if fitz is not None:
        try:
            doc = fitz.open(pdf_path)
            if doc.page_count < 1:
                return ""
            text = doc.load_page(0).get_text("text") or ""
            doc.close()
            return text
        except Exception:
            return ""
    if PdfReader is not None:
        try:
            reader = PdfReader(str(pdf_path))
            return reader.pages[0].extract_text(extraction_mode="layout") or ""
        except Exception:
            return ""
    return ""


def _extract_company_from_fitz_blocks(pdf_path: Path) -> str:
    if fitz is None:
        return ""
    try:
        doc = fitz.open(pdf_path)
        if doc.page_count < 1:
            return ""
        page = doc.load_page(0)
        blocks = page.get_text("blocks") or []
        doc.close()
    except Exception:
        return ""
    blocks.sort(key=lambda b: (b[1], b[0]))
    for x0, y0, x1, y1, text, *_ in blocks:
        if y0 > 220:
            continue
        if not text:
            continue
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        for line in lines:
            lowered = normalize_pdf_text(line)
            if "recibo de" in lowered:
                idx = lowered.find("recibo de")
                if idx > 0:
                    left = line[:idx].strip()
                    if left and not _is_blacklisted_line(left):
                        return _extract_company_from_line(left)
                continue
            if _is_blacklisted_line(line):
                continue
            if (any(ch.isdigit() for ch in line) and not _has_company_token(line)) or CNPJ_RE.search(line):
                continue
            return _extract_company_from_line(line)
    return ""


def extract_first_line(pdf_path: Path) -> str:
    candidate = _extract_company_from_fitz_blocks(pdf_path)
    if candidate:
        return candidate
    text = _read_first_page_text(pdf_path)
    if not text:
        return ""
    try:
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        cutoff_idx = len(lines)
        for idx, line in enumerate(lines):
            lowered = normalize_pdf_text(line)
            if "codigo" in lowered and "nome" in lowered:
                cutoff_idx = idx
                break
            if "descricao" in lowered or "descrição" in lowered:
                cutoff_idx = idx
                break
            if "sal. contr. inss" in lowered:
                cutoff_idx = idx
                break
        if cutoff_idx < len(lines):
            lines = lines[:cutoff_idx]
        if lines:
            first = lines[0]
            lowered = normalize_pdf_text(first)
            if "recibo de" in lowered:
                idx = lowered.find("recibo de")
                if idx > 0:
                    first = first[:idx].strip()
                else:
                    first = ""
            lowered = normalize_pdf_text(first)
            if first and not _is_blacklisted_line(first) and lowered not in ("nome", "nome:") and len(first) >= 5:
                return _extract_company_from_line(first)
        for line in lines:
            lowered = normalize_pdf_text(line)
            if "recibo de" in lowered:
                idx = lowered.find("recibo de")
                if idx > 0:
                    left = line[:idx].strip()
                    if left and not _is_blacklisted_line(left) and not any(ch.isdigit() for ch in left):
                        return left
        for idx, line in enumerate(lines):
            lowered = normalize_pdf_text(line)
            if "empresa" in lowered or "razao social" in lowered:
                if ":" in line:
                    candidate = _extract_company_from_line(line)
                    if candidate and not _looks_like_person(candidate):
                        return candidate
                for nxt in lines[idx + 1:]:
                    if _is_blacklisted_line(nxt):
                        continue
                    if any(ch.isdigit() for ch in nxt):
                        continue
                    if _is_employee_marker(nxt) or CPF_RE.search(nxt):
                        continue
                    if _looks_like_person(nxt):
                        continue
                    return _extract_company_from_line(nxt)
        for idx, line in enumerate(lines):
            if not CNPJ_RE.search(line):
                continue
            for offset in (-1, 1, -2, 2):
                pos = idx + offset
                if pos < 0 or pos >= len(lines):
                    continue
                candidate = lines[pos]
                if _is_blacklisted_line(candidate):
                    continue
                if any(ch.isdigit() for ch in candidate):
                    continue
                if _is_employee_marker(candidate) or CPF_RE.search(candidate):
                    continue
                if _looks_like_person(candidate):
                    continue
                return _extract_company_from_line(candidate)
        best = ""
        best_score = 0
        for line in lines:
            if _is_blacklisted_line(line):
                continue
            if any(ch.isdigit() for ch in line):
                continue
            if _is_employee_marker(line) or CPF_RE.search(line):
                continue
            if _looks_like_person(line):
                continue
            letters = sum(ch.isalpha() for ch in line)
            spaces = sum(ch.isspace() for ch in line)
            if letters < 6 or len(line) < 8:
                continue
            upper = sum(ch.isupper() for ch in line if ch.isalpha())
            upper_ratio = (upper / letters) if letters else 0
            score = letters + spaces + (10 if upper_ratio > 0.6 else 0)
            if _has_company_token(line):
                score += 20
            if score > best_score:
                best_score = score
                best = _extract_company_from_line(line)
        if best:
            return best
    except Exception:
        return ""
    return ""


def get_next_print_index(raw_dir: Path) -> int:
    max_idx = 0
    for pdf in raw_dir.glob("*.pdf"):
        stem = pdf.stem.strip()
        if stem.isdigit():
            max_idx = max(max_idx, int(stem))
    return max_idx + 1


def find_ghostscript() -> str:
    for name in ("gswin64c.exe", "gswin32c.exe", "gswin64c", "gswin32c"):
        path = shutil.which(name)
        if path:
            return path
    gs_root = r"C:\Program Files\gs"
    if os.path.isdir(gs_root):
        candidates = []
        for entry in os.listdir(gs_root):
            bin_path = os.path.join(gs_root, entry, "bin", "gswin64c.exe")
            if os.path.isfile(bin_path):
                candidates.append(bin_path)
        if candidates:
            candidates.sort(reverse=True)
            return candidates[0]
    return ""


def convert_to_pdfa(input_pdf: Path) -> Path:
    gs = find_ghostscript()
    if not gs:
        log("[pdfa] Ghostscript not found; skipping PDF/A conversion.")
        return input_pdf
    out_pdf = input_pdf.with_name(input_pdf.stem + "_pdfa.pdf")
    args = [
        gs,
        "-dPDFA=1",
        "-dBATCH",
        "-dNOPAUSE",
        "-sDEVICE=pdfwrite",
        "-sProcessColorModel=DeviceRGB",
        "-sColorConversionStrategy=RGB",
        "-dEmbedAllFonts=true",
        "-dSubsetFonts=true",
        "-dPDFACompatibilityPolicy=1",
        f"-sOutputFile={out_pdf}",
        str(input_pdf),
    ]
    result = subprocess.run(args, capture_output=True, text=True)
    if result.returncode != 0 or not out_pdf.exists():
        log(f"[pdfa] conversion failed: {result.stderr.strip()}")
        return input_pdf
    try:
        out_pdf.replace(input_pdf)
    except Exception:
        pass
    return input_pdf


def extract_company_code_from_path(pdf_path: Path) -> str:
    # Uses the 4-digit folder in the path (e.g., 0137 -> 137).
    parts = [p.name for p in pdf_path.parents]
    for name in parts:
        if len(name) == 4 and name.isdigit():
            return str(int(name))
    return ""


def extract_company_code_from_name(file_name: str) -> str:
    # Matches " - 123 - Pro Labore.pdf" patterns in final output names.
    name = file_name or ""
    parts = name.split(" - ")
    for part in parts:
        if part.isdigit():
            return str(int(part))
    return ""


def move_pdf_to_output(pdf_path: Path, output_root: str = OUTPUT_ROOT, year: str | None = None, month: str | None = None):
    now = datetime.now()
    year = year or str(now.year)
    month = month or f"{now.month:02d}"
    first_line = extract_first_line(pdf_path)
    company = sanitize_filename(first_line) or sanitize_filename(pdf_path.stem)
    code = extract_company_code_from_path(pdf_path)
    if not code:
        code = extract_company_code_from_name(pdf_path.name)
    code_suffix = f" - {code}" if code else ""
    file_name = f"{company}{code_suffix} - Pro Labore.pdf"
    dest_dir = Path(output_root) / year / month / company
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_path = dest_dir / file_name
    if dest_path.exists():
        idx = 2
        while True:
            alt = dest_dir / f"{company}{code_suffix} - Pro Labore ({idx}).pdf"
            if not alt.exists():
                dest_path = alt
                break
            idx += 1
    shutil.copy2(str(pdf_path), str(dest_path))
    log(f"[pdf] saved: {dest_path}")


def process_pdf_file(pdf_path: Path):
    if not wait_file_stable(pdf_path):
        log(f"[pdf] not stable yet: {pdf_path}")
        return
    convert_to_pdfa(pdf_path)
    move_pdf_to_output(pdf_path)


def reprocess_output_folder(output_root: str = OUTPUT_ROOT):
    root = Path(output_root)
    if not root.exists():
        log(f"[reprocess] output root not found: {root}")
        return
    pdfs = list(root.rglob("*.pdf"))
    for pdf in pdfs:
        try:
            rel_parts = pdf.relative_to(root).parts
        except Exception:
            rel_parts = []
        year = month = None
        if len(rel_parts) >= 2 and rel_parts[0].isdigit() and rel_parts[1].isdigit():
            year = rel_parts[0]
            month = rel_parts[1]
        move_pdf_to_output(pdf, output_root=output_root, year=year, month=month)


def snapshot_dir_state(raw_dir: Path) -> dict:
    state = {}
    for p in raw_dir.rglob("*"):
        if p.is_dir():
            continue
        try:
            stat = p.stat()
        except Exception:
            continue
        state[str(p)] = (stat.st_mtime, stat.st_size)
    return state


def process_pdfs_until_idle(raw_dir: Path, idle_s: int = IDLE_SECONDS):
    log(f"[watch] monitoring: {raw_dir}")
    seen = set()
    last_activity = time.time()
    last_state = snapshot_dir_state(raw_dir)
    had_activity = False
    print_state = {"next_idx": get_next_print_index(raw_dir)}
    last_pdf_scan = 0.0
    while True:
        handle_print_offers(raw_dir, print_state)
        if (time.time() - last_pdf_scan) >= POLL_SECONDS:
            pdfs = sorted(raw_dir.rglob("*.pdf"))
            for pdf in pdfs:
                if pdf in seen:
                    continue
                process_pdf_file(pdf)
                seen.add(pdf)
                had_activity = True
            last_pdf_scan = time.time()
        current_state = snapshot_dir_state(raw_dir)
        if current_state != last_state:
            last_activity = time.time()
            last_state = current_state
            had_activity = True
        if had_activity and (time.time() - last_activity) >= idle_s:
            break
        time.sleep(PRINT_OFFER_POLL)
    handle_print_offers(raw_dir, print_state)
    pdfs = sorted(raw_dir.rglob("*.pdf"))
    for pdf in pdfs:
        if pdf in seen:
            continue
        process_pdf_file(pdf)


def finalize_after_pdfs():
    click_abs(FINAL_OK_COORD)
    time.sleep(FINAL_BETWEEN_WAIT_S)
    click_abs(FINAL_CLOSE_COORD)
    time.sleep(FINAL_CLOSE_WAIT_S)


# =========================
# UI automation
# =========================

def attach_main_window(title_contains: str):
    desk = Desktop(backend="uia")
    needle = (title_contains or "").lower()
    for w in desk.windows():
        title = (w.window_text() or "").lower()
        if needle and needle in title:
            return w
    return None


def focus_main_window(title_contains: str, timeout_s: float = 10.0) -> bool:
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        win = attach_main_window(title_contains)
        if not win:
            time.sleep(0.2)
            continue
        try:
            if hasattr(win, "is_minimized") and win.is_minimized():
                win.restore()
        except Exception:
            pass
        try:
            if hasattr(win, "is_maximized") and not win.is_maximized():
                win.maximize()
        except Exception:
            pass
        try:
            win.set_focus()
        except Exception:
            pass
        time.sleep(0.2)
        try:
            if win.is_active():
                return True
        except Exception:
            pass
    return False


def capture_click_coordinates(label: str, timeout_s: float = 30.0):
    print(f"\n>>> Press ENTER and click on '{label}' to capture coordinates.")
    input("Press ENTER to arm capture...")
    try:
        while ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000:
            time.sleep(0.01)
        t0 = time.time()
        while (time.time() - t0) < timeout_s:
            if ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000:
                pt = wintypes.POINT()
                if ctypes.windll.user32.GetCursorPos(ctypes.byref(pt)):
                    print(f"[coord] {label}: x={pt.x} y={pt.y}")
                    return (pt.x, pt.y)
                raise OSError("GetCursorPos failed")
            time.sleep(0.01)
    except Exception as exc:
        print(f"[coord] failed for {label}: {exc}")
    return None


def normalize_text(text: str) -> str:
    return (text or "").strip().lower()


def find_window_by_title(keywords):
    desk = Desktop(backend="uia")
    for w in desk.windows():
        title = normalize_text(w.window_text())
        if not title:
            continue
        if any(k in title for k in keywords):
            return w
    return None


def find_button_by_keywords(root, keywords):
    buttons = root.descendants(control_type="Button")
    for b in buttons:
        name = normalize_text(b.window_text() or b.element_info.name)
        if any(k in name for k in keywords):
            return b
    return None


def set_printer_to_pdf(root):
    combos = root.descendants(control_type="ComboBox")
    if not combos:
        return False
    target = "Microsoft Print to PDF"
    for cb in combos:
        try:
            cb.set_focus()
            time.sleep(0.1)
            send_keys("^a")
            send_keys(target)
            send_keys("{ENTER}")
            time.sleep(0.2)
            return True
        except Exception:
            continue
    return False


def ensure_print_to_pdf(root) -> bool:
    combos = root.descendants(control_type="ComboBox")
    if not combos:
        return False
    target = "microsoft print to pdf"
    for cb in combos:
        try:
            current = normalize_text(cb.window_text() or cb.get_value())
        except Exception:
            current = ""
        if target in current:
            return True
    return set_printer_to_pdf(root)


def log_print_issue(msg: str, raw_dir: Path | None = None):
    try:
        base = raw_dir or Path(get_raw_dir())
        log_path = base / "print_errors.log"
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def save_print_output(raw_dir: Path, index: int) -> bool:
    full_path = raw_dir / f"{index}.pdf"
    desk = Desktop(backend="uia")
    t0 = time.time()
    while (time.time() - t0) < 30.0:
        wins = desk.windows()
        for w in wins:
            title = normalize_text(w.window_text())
            if "salvar" not in title and "save" not in title:
                continue
            try:
                w.set_focus()
            except Exception:
                pass
            time.sleep(0.15)
            send_keys("%d")
            time.sleep(0.15)
            send_keys(str(full_path.parent))
            send_keys("{ENTER}")
            time.sleep(0.4)
            try:
                name_edit = w.child_window(auto_id="1001")
                if name_edit:
                    name_edit.set_focus()
                    time.sleep(0.1)
                    send_keys("^a")
                    send_keys(full_path.name)
                else:
                    send_keys(full_path.name)
            except Exception:
                send_keys(full_path.name)
            time.sleep(0.15)
            try:
                btn = w.child_window(auto_id="1")
                if btn:
                    btn.click_input()
                else:
                    send_keys("%s")
                    send_keys("{ENTER}")
            except Exception:
                send_keys("%s")
                send_keys("{ENTER}")
            return True
        time.sleep(0.2)
    return False


def handle_print_offers(raw_dir: Path, print_state: dict):
    print_win = find_window_by_title(["imprimir", "print"])
    if not print_win:
        return
    try:
        print_win.set_focus()
    except Exception:
        pass
    time.sleep(0.2)
    if not ensure_print_to_pdf(print_win):
        time.sleep(0.3)
        if not ensure_print_to_pdf(print_win):
            log("[print] unable to confirm Microsoft Print to PDF; skipping print.")
            log_print_issue("Failed to select Microsoft Print to PDF.", raw_dir=raw_dir)
            return
    btn = find_button_by_keywords(print_win, ["imprimir", "print", "ok"])
    if btn:
        try:
            btn.click_input()
        except Exception:
            send_keys("{ENTER}")
    else:
        send_keys("{ENTER}")
    time.sleep(0.4)
    idx = print_state.get("next_idx", 1)
    if save_print_output(raw_dir, idx):
        print_state["next_idx"] = idx + 1


def run_iob_flow(main_win, coords: dict, skip_processar: bool = False):
    main_win.set_focus()
    time.sleep(0.25)

    send_keys("%m")
    time.sleep(0.2)
    send_keys("e")
    time.sleep(0.5)
    click_abs((1568, 310))
    time.sleep(0.4)

    send_keys(PROCESSAMENTO_VALUE)
    send_keys("{ENTER}")
    time.sleep(0.5)

    send_keys(MES_ANO_VALUE)
    send_keys("{ENTER}")
    time.sleep(0.5)

    send_keys(GERAR_POR_VALUE)
    send_keys("{ENTER}")
    time.sleep(0.5)

    send_keys(GRUPO_VALUE)
    send_keys("{ENTER}")
    time.sleep(3.0)

    if coords.get("ctx_calc_folha"):
        context_menu_sequence(coords["ctx_calc_folha"], "DDRE", delay_s=0.2, post_delay_s=0.5)
        context_menu_sequence(coords["ctx_calc_folha"], "DDRE", delay_s=0.2, post_delay_s=0.5)
    if coords.get("ctx_apuracao"):
        context_menu_sequence(coords["ctx_apuracao"], "DDRE", delay_s=0.2, post_delay_s=0.5)
        context_menu_sequence(coords["ctx_apuracao"], "DDRE", delay_s=0.2, post_delay_s=0.5)
    if coords.get("ctx_holerite"):
        context_menu_sequence(coords["ctx_holerite"], "DDRE", delay_s=0.2, post_delay_s=0.5)
        context_menu_sequence(coords["ctx_holerite"], "DDRE", delay_s=0.2, post_delay_s=0.5)
    if coords.get("ctx_extra"):
        # Desmarcar coluna "Rel. Folha".
        context_menu_sequence(coords["ctx_extra"], "DDRDE", delay_s=0.2, post_delay_s=0.5)
        context_menu_sequence(coords["ctx_extra"], "DDRDE", delay_s=0.2, post_delay_s=0.5)

    time.sleep(UI_CLICK_GAP_S)
    click_abs(coords["menu_parametros"])
    time.sleep(UI_CLICK_GAP_S)
    click_abs(coords["menu_relatorio_mensal"])
    time.sleep(UI_CLICK_GAP_S)

    click_abs(coords["btn_diretorio"])
    time.sleep(0.6)
    # Garante foco na barra de enderecos da janela "Selecionar pasta".
    send_keys("%d")
    time.sleep(0.2)
    send_keys(get_raw_dir())
    send_keys("{ENTER}")
    time.sleep(0.4)
    send_keys("{ENTER}")
    time.sleep(0.4)
    sel_win = find_window_by_title(["seleção de diretório", "selecao de diretorio"])
    if sel_win:
        btn = find_button_by_keywords(sel_win, ["selecionar pasta"])
        if btn:
            try:
                btn.click_input()
            except Exception:
                pass
            time.sleep(0.6)
        elif coords.get("btn_selecionar_pasta"):
            click_abs(coords["btn_selecionar_pasta"])
            time.sleep(0.6)

    if not skip_processar:
        click_abs(coords["btn_processar"])


def guided_capture_coords(path: str):
    data = {}
    for key in COORDS.keys():
        label = key.replace("_", " ")
        pos = capture_click_coordinates(label)
        if pos:
            data[key] = {"x": pos[0], "y": pos[1]}
    save_coords(path, data)
    log(f"[coord] saved: {path}")


def normalize_coords(data: dict) -> dict:
    out = {}
    for key in COORDS.keys():
        value = data.get(key)
        if isinstance(value, dict) and "x" in value and "y" in value:
            out[key] = (int(value["x"]), int(value["y"]))
        elif isinstance(value, (list, tuple)) and len(value) == 2:
            out[key] = (int(value[0]), int(value[1]))
        else:
            out[key] = None
    return out


def validate_coords(coords: dict):
    required_keys = [
        "ctx_calc_folha",
        "ctx_apuracao",
        "ctx_holerite",
        "ctx_extra",
        "menu_parametros",
        "menu_relatorio_mensal",
        "btn_diretorio",
        "btn_selecionar_pasta",
        "btn_processar",
    ]
    missing = [k for k in required_keys if not coords.get(k)]
    if missing:
        raise SystemExit(f"Missing coordinates: {', '.join(missing)}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--title-contains", default=ATTACH_TITLE_CONTAINS)
    ap.add_argument("--coords-file", default=COORDS_PATH)
    ap.add_argument("--capture-coords", action="store_true")
    ap.add_argument("--skip-ui", action="store_true", help="Skip UI automation and only post-process PDFs.")
    ap.add_argument("--post-process-only", action="store_true", help="Only convert/rename/move PDFs in the raw folder.")
    ap.add_argument("--skip-processar", action="store_true", help="Run the UI flow but skip clicking Processar.")
    ap.add_argument("--reprocess-output", action="store_true", help="Reprocess PDFs in the final output folder.")
    ap.add_argument("--idle-seconds", type=int, default=IDLE_SECONDS)
    args = ap.parse_args()

    if args.capture_coords:
        guided_capture_coords(args.coords_file)
        return

    raw_dir = Path(get_raw_dir())
    raw_dir.mkdir(parents=True, exist_ok=True)

    if args.post_process_only:
        process_pdfs_until_idle(raw_dir, idle_s=args.idle_seconds)
        return
    if args.reprocess_output:
        reprocess_output_folder()
        return

    if not args.skip_ui:
        main_win = attach_main_window(args.title_contains)
        if not main_win:
            raise SystemExit("Main IOB window not found. Adjust --title-contains.")
        if not focus_main_window(args.title_contains):
            raise SystemExit("Unable to focus Folha de Pagamento window.")
        coords_data = load_coords(args.coords_file)
        coords = normalize_coords(coords_data)
        validate_coords(coords)
        run_iob_flow(main_win, coords, skip_processar=args.skip_processar)
        if args.skip_processar:
            log("[run] --skip-processar ativo: fluxo encerrado sem pos-processamento de PDFs.")
            return

    process_pdfs_until_idle(raw_dir, idle_s=args.idle_seconds)
    finalize_after_pdfs()

if __name__ == "__main__":
    main()
