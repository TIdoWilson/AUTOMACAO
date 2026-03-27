import argparse
import json
import os
import re
import shutil
import subprocess
import time
import unicodedata
from datetime import datetime
from pathlib import Path

try:
    from pywinauto import Desktop, mouse
    from pywinauto.keyboard import send_keys as send_keys_raw
except Exception as exc:
    raise SystemExit("pywinauto is required for UI automation") from exc

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None


# =========================
# CONFIG
# =========================

ATTACH_TITLE_CONTAINS = "folha de pagamento"
PROCESSAMENTO_VALUE = "2"
GERAR_POR_VALUE = "1"
GRUPOS = ["6", "13", "14"]

BASE_DIR = Path(__file__).resolve().parent
RAW_ROOT = BASE_DIR / "FOLHAS"
RAW_INBOX_NAME = "entrada"
RAW_ARCHIVE_NAME = "brutos"
OUTPUT_ROOT = Path(r"W:\\DOCUMENTOS ESCRITORIO\\RH\\AUTOMATIZADO\\1ª PARTE")

COORDS_PATH = Path(__file__).with_suffix(".json")

KEY_PAUSE = 0.04
CLICK_DELAY = 1.2
FIELD_DELAY = 1.0
POLL_SECONDS = 120
PRINT_OFFER_POLL = 1.5
IDLE_SECONDS = 180
GROUP_CHANGE_DELAY = 6
DIR_CLICK_OFFSET_X = -30
DIR_CHECKBOX_COORD = (598, 783)


COORDS = {
    "field_data_sistema": None,
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

OK_COMPLETO_COORD = (1079, 602)
CLOSE_TAB_COORD = (1475, 165)
CLICK_Y_OFFSET = -2


# =========================
# UTIL
# =========================

def log(msg: str):
    print(msg, flush=True)


def load_coords(path: Path) -> dict:
    if not path or not path.exists():
        return {}
    with path.open("r", encoding="utf-8-sig") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        return {}
    return data


def save_coords(path: Path, data: dict):
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def send_keys(text: str, pause: float = KEY_PAUSE):
    send_keys_raw(text, pause=pause, with_spaces=True)


def click_abs(pos):
    x, y = pos
    mouse.click(button="left", coords=(x, y + CLICK_Y_OFFSET))
    time.sleep(CLICK_DELAY)


def click_abs_retry(pos, retries: int = 2, delay_s: float = 0.4):
    for _ in range(max(1, retries)):
        click_abs(pos)
        time.sleep(delay_s)


def replace_text_at(coord, text: str, clear_first: bool = False):
    click_abs(coord)
    time.sleep(FIELD_DELAY)
    if clear_first:
        time.sleep(5.0)
        send_keys("{BACKSPACE}" * 8)
        send_keys("{DELETE}" * 8)
        time.sleep(0.2)
    send_keys(text)
    time.sleep(FIELD_DELAY)


def set_diretorio_sem_explorer(coord, path: Path):
    x, y = coord
    click_abs((x + DIR_CLICK_OFFSET_X, y))
    time.sleep(FIELD_DELAY)
    send_keys("^a")
    send_keys(str(path))
    send_keys("{ENTER}")
    time.sleep(FIELD_DELAY)

def set_data_sistema(coords: dict, date_value: str):
    click_abs(coords["field_data_sistema"])
    time.sleep(FIELD_DELAY)
    send_keys("{DELETE}" * 8)
    send_keys("{BACKSPACE}" * 8)
    send_keys(date_value)
    send_keys("{ENTER}")
    send_keys("{ENTER}")
    time.sleep(FIELD_DELAY)


def context_menu_sequence(coord, moves: str, delay_s: float = 1.0, post_delay_s: float = 1.0):
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

def normalize_pdf_text(text: str) -> str:
    raw = (text or "").strip().lower()
    return "".join(ch for ch in unicodedata.normalize("NFKD", raw) if not unicodedata.combining(ch))


def _extract_company_from_line(line: str) -> str:
    text = (line or "").strip()
    if not text:
        return ""
    lowered = text.lower()
    for token in ("empresa:", "razao social:", "razÃ£o social:"):
        if token in lowered:
            parts = text.split(":", 1)
            if len(parts) == 2:
                return parts[1].strip()
    return text


def _is_blacklisted_line(line: str) -> bool:
    lowered = normalize_pdf_text(line)
    return any(phrase in lowered for phrase in BLACKLIST_PHRASES)


BLACKLIST_PHRASES = [
    "recibo",
    "folha",
    "competencia",
    "referente",
    "cpf",
    "cnpj",
    "matricula",
    "cod.",
    "codigo",
    "total de vencimentos",
    "total vencimentos",
    "relatorio",
    "processamento",
    "trial mode",
    "click here for more information",
    "codigo nome do funcionario",
    "cbo emp",
    "local depto",
    "setor",
    "secao",
]


def extract_company_from_pypdf(pdf_path: Path) -> str:
    if PdfReader is None:
        return ""
    try:
        reader = PdfReader(str(pdf_path))
        text = reader.pages[0].extract_text(extraction_mode="layout") or ""
    except Exception:
        return ""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for line in lines:
        lowered = normalize_pdf_text(line)
        if "trial mode" in lowered:
            continue
        if "recibo de" in lowered:
            idx = lowered.find("recibo de")
            if idx > 0:
                left = line[:idx].strip()
                if left and not _is_blacklisted_line(left) and not any(ch.isdigit() for ch in left):
                    return left
    for line in lines:
        if _is_blacklisted_line(line):
            continue
        if any(ch.isdigit() for ch in line):
            continue
        if len(line) < 8:
            continue
        return line
    return ""


def extract_first_line(pdf_path: Path) -> str:
    if PdfReader is None:
        raise SystemExit("pypdf is required for PDF name extraction. Install with: pip install pypdf")
    return extract_company_from_pypdf(pdf_path)


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
    # Matches " - 123 - Folha 1ª parte.pdf" patterns in final output names.
    name = file_name or ""
    parts = name.split(" - ")
    for part in parts:
        if part.isdigit():
            return str(int(part))
    return ""


def move_pdf_to_output(pdf_path: Path, grupo: str, output_root: Path = OUTPUT_ROOT, year: str | None = None, month: str | None = None):
    now = datetime.now()
    year = year or str(now.year)
    month = month or f"{now.month:02d}"
    first_line = extract_first_line(pdf_path)
    company = sanitize_filename(first_line) or sanitize_filename(pdf_path.stem)
    code = extract_company_code_from_path(pdf_path)
    if not code:
        code = extract_company_code_from_name(pdf_path.name)
    code_suffix = f" - {code}" if code else ""
    if not company:
        company = "SEM_NOME"
    file_name = f"{company}{code_suffix} - Folha 1ª parte.pdf"
    dest_dir = output_root / year / month / str(grupo) / company
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_path = dest_dir / file_name
    if dest_path.exists():
        idx = 2
        while True:
            alt = dest_dir / f"{company}{code_suffix} - Folha 1ª parte ({idx}).pdf"
            if not alt.exists():
                dest_path = alt
                break
            idx += 1
    shutil.move(str(pdf_path), str(dest_path))
    log(f"[pdf] saved: {dest_path}")


def resolve_year_month(year: str | None = None, month: str | None = None) -> tuple[str, str]:
    now = datetime.now()
    year = year or str(now.year)
    month = month or f"{now.month:02d}"
    return year, month


def get_raw_dirs(grupo: str, year: str | None = None, month: str | None = None) -> tuple[Path, Path]:
    year, month = resolve_year_month(year, month)
    base = RAW_ROOT / year / month / str(grupo)
    inbox = base / RAW_INBOX_NAME
    archive = base / RAW_ARCHIVE_NAME
    inbox.mkdir(parents=True, exist_ok=True)
    archive.mkdir(parents=True, exist_ok=True)
    return inbox, archive


def ensure_raw_structure(grupos: list[str], year: str | None = None, month: str | None = None):
    for grupo in grupos:
        get_raw_dirs(grupo, year=year, month=month)


def store_raw_copy(pdf_path: Path, archive_dir: Path) -> Path:
    archive_dir.mkdir(parents=True, exist_ok=True)
    dest = archive_dir / pdf_path.name
    if dest.exists():
        stem = pdf_path.stem
        suffix = pdf_path.suffix
        idx = 2
        while True:
            candidate = archive_dir / f"{stem} ({idx}){suffix}"
            if not candidate.exists():
                dest = candidate
                break
            idx += 1
    shutil.copy2(pdf_path, dest)
    return dest


def process_pdf_file(pdf_path: Path, grupo: str, archive_dir: Path | None = None):
    if not wait_file_stable(pdf_path):
        log(f"[pdf] not stable yet: {pdf_path}")
        return
    if archive_dir:
        store_raw_copy(pdf_path, archive_dir)
    convert_to_pdfa(pdf_path)
    move_pdf_to_output(pdf_path, grupo=grupo)


def process_pdf_file_with_target(
    pdf_path: Path,
    grupo: str,
    year: str,
    month: str,
    archive_dir: Path | None = None,
):
    if not wait_file_stable(pdf_path):
        log(f"[pdf] not stable yet: {pdf_path}")
        return
    if archive_dir:
        store_raw_copy(pdf_path, archive_dir)
    convert_to_pdfa(pdf_path)
    move_pdf_to_output(pdf_path, grupo=grupo, year=year, month=month)


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


def process_pdfs_until_idle(
    raw_dir: Path,
    grupo: str,
    idle_s: int = IDLE_SECONDS,
    archive_dir: Path | None = None,
):
    log(f"[watch] monitoring: {raw_dir}")
    seen = set()
    last_activity = time.time()
    last_state = snapshot_dir_state(raw_dir)
    had_activity = False
    last_pdf_scan = 0.0
    while True:
        handle_print_offers(raw_dir)
        if (time.time() - last_pdf_scan) >= POLL_SECONDS:
            pdfs = sorted(raw_dir.rglob("*.pdf"))
            for pdf in pdfs:
                if pdf in seen:
                    continue
                process_pdf_file(pdf, grupo=grupo, archive_dir=archive_dir)
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
    handle_print_offers(raw_dir)
    pdfs = sorted(raw_dir.rglob("*.pdf"))
    for pdf in pdfs:
        if pdf in seen:
            continue
        process_pdf_file(pdf, grupo=grupo, archive_dir=archive_dir)


def infer_group_year_month_from_path(pdf_path: Path, grupos: list[str]) -> tuple[str | None, str | None, str | None]:
    grupo = None
    year = None
    month = None

    for part in pdf_path.parts:
        token = str(part).strip()
        if token in grupos:
            grupo = token
            break
        m_grupo = re.match(r"^GRUPO\s*0?(\d+)$", token, re.IGNORECASE)
        if m_grupo:
            g = str(int(m_grupo.group(1)))
            if g in grupos:
                grupo = g
                break

    for part in pdf_path.parts:
        m_comp = re.match(r"^(20\d{2})(0[1-9]|1[0-2])$", str(part).strip())
        if m_comp:
            year = m_comp.group(1)
            month = m_comp.group(2)

    return grupo, year, month


def process_pdfs_from_source_root(source_root: Path, grupos: list[str]):
    if not source_root.exists():
        raise SystemExit(f"Pasta de origem nao encontrada: {source_root}")

    pdfs = sorted(source_root.rglob("*.pdf"))
    if not pdfs:
        log(f"[info] nenhum PDF encontrado em: {source_root}")
        return

    processed = 0
    skipped = 0
    for pdf in pdfs:
        grupo, year, month = infer_group_year_month_from_path(pdf, grupos)
        if not grupo or not year or not month:
            skipped += 1
            log(f"[skip] caminho sem grupo/competencia identificavel: {pdf}")
            continue
        _, archive_dir = get_raw_dirs(grupo, year=year, month=month)
        process_pdf_file_with_target(
            pdf,
            grupo=grupo,
            year=year,
            month=month,
            archive_dir=archive_dir,
        )
        processed += 1

    log(f"[ok] PDFs processados: {processed}")
    if skipped:
        log(f"[info] PDFs ignorados (sem grupo/competencia): {skipped}")


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


def normalize_text(text: str) -> str:
    raw = (text or "").strip().lower()
    # Normalize to ASCII-ish text for window title matching.
    return "".join(ch for ch in unicodedata.normalize("NFKD", raw) if not unicodedata.combining(ch))


def find_window_by_title(keywords):
    desk = Desktop(backend="uia")
    for w in desk.windows():
        title = normalize_text(w.window_text())
        if not title:
            continue
        if any(k in title for k in keywords):
            return w
    return None


def wait_window_by_title(keywords, timeout_s: float = 8.0):
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        win = find_window_by_title(keywords)
        if win:
            return win
        time.sleep(0.2)
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
        base = raw_dir or RAW_ROOT
        log_path = base / "print_errors.log"
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def save_print_output(raw_dir: Path) -> bool:
    full_path = raw_dir / f"{int(time.time())}.pdf"
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


def handle_print_offers(raw_dir: Path):
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
    save_print_output(raw_dir)


def close_completion_dialog(timeout_s: float = 20.0, click_coord: bool = False, post_click_delay_s: float = 3.0) -> bool:
    desk = Desktop(backend="uia")
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        for w in desk.windows():
            title = normalize_text(w.window_text())
            if "aviso do sistema" not in title:
                continue
            try:
                text = normalize_text(" ".join(w.texts()))
            except Exception:
                text = ""
            if "processamento" not in text or "sucesso" not in text:
                continue
            try:
                w.set_focus()
            except Exception:
                pass
            if click_coord:
                time.sleep(post_click_delay_s)
                try:
                    click_abs(OK_COMPLETO_COORD)
                except Exception:
                    pass
            return True
        time.sleep(0.2)
    return False


def run_iob_flow(
    main_win,
    coords: dict,
    grupo: str,
    raw_dir: Path,
    base_setup: bool = True,
    refresh_options: bool = True,
    skip_processar: bool = False,
):
    main_win.set_focus()
    time.sleep(0.25)

    if base_setup:
        data_sistema = datetime.now().strftime("%d/%m/%Y")
        set_data_sistema(coords, data_sistema)

        send_keys("%m")
        time.sleep(0.2)
        send_keys("e")
        time.sleep(0.5)
        click_abs((1568, 310))
        time.sleep(0.4)

        replace_text_at(coords["field_processamento"], PROCESSAMENTO_VALUE)

        mes_ano = datetime.now().strftime("%m/%Y")
        replace_text_at(coords["field_mes_ano"], mes_ano, clear_first=True)
        send_keys("{ENTER}")
        time.sleep(0.5)

        replace_text_at(coords["field_gerar_por"], GERAR_POR_VALUE)
        send_keys("{ENTER}")
        time.sleep(0.5)

    replace_text_at(coords["field_grupo"], str(grupo), clear_first=True)
    send_keys("{ENTER}")
    time.sleep(GROUP_CHANGE_DELAY)

    if refresh_options:
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
            context_menu_sequence(coords["ctx_extra"], "DDRDE", delay_s=0.2, post_delay_s=0.5)
            context_menu_sequence(coords["ctx_extra"], "DDRDE", delay_s=0.2, post_delay_s=0.5)

        click_abs_retry(coords["menu_parametros"])
        click_abs_retry(coords["menu_relatorio_mensal"])

        if str(grupo) in ("13", "14"):
            click_abs(DIR_CHECKBOX_COORD)
            time.sleep(0.2)
        set_diretorio_sem_explorer(coords["btn_diretorio"], raw_dir)

    if skip_processar:
        log("[info] pular processar (modo teste)")
        return
    click_abs(coords["btn_processar"])


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
    missing = [k for k, v in coords.items() if not v]
    if missing:
        raise SystemExit(f"Missing coordinates: {', '.join(missing)}")


def parse_grupos(value: str | None) -> list[str]:
    if not value:
        return GRUPOS
    parts = re.split(r"[;, ]+", value.strip())
    grupos = []
    for part in parts:
        if not part:
            continue
        if part not in GRUPOS:
            raise SystemExit(f"Grupo invalido: {part}. Validos: {', '.join(GRUPOS)}")
        if part not in grupos:
            grupos.append(part)
    if not grupos:
        raise SystemExit(f"Nenhum grupo informado. Validos: {', '.join(GRUPOS)}")
    return grupos


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--title-contains", default=ATTACH_TITLE_CONTAINS)
    ap.add_argument("--coords-file", default=str(COORDS_PATH))
    ap.add_argument("--skip-ui", action="store_true", help="Skip UI automation and only post-process PDFs.")
    ap.add_argument(
        "--skip-processar",
        action="store_true",
        help="Nao clica em Processar (modo teste).",
    )
    ap.add_argument(
        "--skip-pdfs",
        action="store_true",
        help="Nao processa/organiza PDFs (modo teste).",
    )
    ap.add_argument("--idle-seconds", type=int, default=IDLE_SECONDS)
    ap.add_argument(
        "--organize-from-root",
        default=None,
        help="Organiza PDFs existentes a partir de uma raiz (ex: ...\\4 - 1ª PARTE\\FOLHAS).",
    )
    ap.add_argument(
        "--grupos",
        default=None,
        help="Lista de grupos para rodar (ex: 13,14 ou 13;14).",
    )
    args = ap.parse_args()

    grupos = parse_grupos(args.grupos)
    ensure_raw_structure(grupos)

    if args.organize_from_root:
        process_pdfs_from_source_root(Path(args.organize_from_root), grupos=grupos)
        return

    if args.skip_ui:
        for grupo in grupos:
            raw_inbox, raw_archive = get_raw_dirs(grupo)
            process_pdfs_until_idle(
                raw_inbox,
                grupo=grupo,
                idle_s=args.idle_seconds,
                archive_dir=raw_archive,
            )
        return

    main_win = attach_main_window(args.title_contains)
    if not main_win:
        raise SystemExit("Main IOB window not found. Adjust --title-contains.")
    coords_data = load_coords(Path(args.coords_file))
    coords = normalize_coords(coords_data)
    validate_coords(coords)

    for idx, grupo in enumerate(grupos):
        log(f"[grupo] iniciando grupo {grupo}")
        raw_inbox, raw_archive = get_raw_dirs(grupo)
        run_iob_flow(
            main_win,
            coords,
            grupo=grupo,
            raw_dir=raw_inbox,
            base_setup=True,
            refresh_options=True,
            skip_processar=args.skip_processar,
        )
        if not args.skip_pdfs:
            process_pdfs_until_idle(
                raw_inbox,
                grupo=grupo,
                idle_s=args.idle_seconds,
                archive_dir=raw_archive,
            )
        if idx < len(grupos) - 1:
            click_abs(OK_COMPLETO_COORD)
            time.sleep(10.0)
            click_abs(CLOSE_TAB_COORD)
        try:
            main_win.set_focus()
        except Exception:
            pass


if __name__ == "__main__":
    main()



