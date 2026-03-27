import argparse
import ctypes
import ctypes.wintypes as wintypes
import json
import os
import shutil
import time
from datetime import datetime

try:
    from pywinauto import Desktop, mouse
    from pywinauto.keyboard import send_keys as send_keys_raw
except Exception as exc:
    raise SystemExit("pywinauto is required for UI automation") from exc


# =========================
# CONFIG
# =========================

ATTACH_TITLE_CONTAINS = "folha de pagamento"
PROCESSAMENTO_VALUE = "2"
GERAR_POR_VALUE = "1"
GRUPO_VALUE = "3"
OUTPUT_DIR_TEXT = "W:\\DOCUMENTOS ESCRITORIO\\INSTALACAO SISTEMA\\python\\PROJETO RH LEONARDO\\4 - 1\u00aa PARTE\\Folhas de Pagamento"
OUTPUT_DIR = OUTPUT_DIR_TEXT
AUTOMATIZADO_ROOT = "W:\\DOCUMENTOS ESCRITORIO\\RH\\AUTOMATIZADO\\DOM\u00c9STICAS"
EXCEL_PREP_PATH = "W:\\DOCUMENTOS ESCRITORIO\\INSTALACAO SISTEMA\\python\\PROJETO RH LEONARDO\\4 - 1\u00aa PARTE\\domesticas_preparado.xlsx"
EXCEL_SHEET = "Dom\u00e9sticas"
EXCEL_COL_NUM = "N"
EXCEL_COL_NOME = "Nome"
COORDS_PATH = os.path.join(os.path.dirname(__file__), "coords_domesticas_processamento.json")
PROLABORE_COORDS_PATH = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\1 - Pro Labore\coords_prolabore.json"

KEY_PAUSE = 0.04
CLICK_DELAY = 0.2
FIELD_DELAY = 0.2
UI_CLICK_GAP_S = 1.0
GROUP_WAIT_S = 4.0

MARK_SEQUENCE = "DDRE"
UNMARK_SEQUENCE = "DDRDE"
SEQUENCE_REPEAT = 2


COORDS = {
    "field_processamento": None,
    "field_mes_ano": None,
    "field_gerar_por": None,
    "field_grupo": None,
    "ctx_calc_folha": None,
    "ctx_holerite": None,
    "ctx_apuracao": None,
    "ctx_guias": None,
    "ctx_rel_folha": None,
    "menu_parametros": None,
    "menu_relatorio_mensal": None,
    "btn_diretorio": None,
    "btn_selecionar_pasta": None,
    "btn_processar": None,
}

DEFAULT_COORDS = {
    "field_processamento": (580, 200),
    "field_mes_ano": (580, 220),
    "field_gerar_por": (580, 240),
    "field_grupo": (580, 260),
    "menu_parametros": (518, 293),
    "menu_relatorio_mensal": (575, 322),
    "btn_diretorio": (1035, 810),
    "btn_selecionar_pasta": (1550, 770),
    "btn_processar": (1420, 200),
    "ctx_calc_folha": (824, 322),
    "ctx_holerite": (1034, 321),
    "ctx_apuracao": (930, 322),
    "ctx_guias": (990, 322),
    "ctx_rel_folha": (1087, 321),
}


# =========================
# UTIL
# =========================

def log(msg: str):
    print(msg, flush=True)


def load_coords(path: str) -> dict:
    if not path or not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8-sig") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        return {}
    return data


def save_coords(path: str, data: dict):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def send_keys(text: str, pause: float = KEY_PAUSE):
    send_keys_raw(text, pause=pause, with_spaces=True)


def set_clipboard_text(text: str):
    if text is None:
        return
    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002
    data = str(text)
    h_global = ctypes.windll.kernel32.GlobalAlloc(GMEM_MOVEABLE, (len(data) + 1) * 2)
    if not h_global:
        return
    locked = ctypes.windll.kernel32.GlobalLock(h_global)
    if not locked:
        ctypes.windll.kernel32.GlobalFree(h_global)
        return
    ctypes.memmove(locked, ctypes.create_unicode_buffer(data), (len(data) + 1) * 2)
    ctypes.windll.kernel32.GlobalUnlock(h_global)
    ctypes.windll.user32.OpenClipboard(0)
    ctypes.windll.user32.EmptyClipboard()
    ctypes.windll.user32.SetClipboardData(CF_UNICODETEXT, h_global)
    ctypes.windll.user32.CloseClipboard()


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


def replace_text_at_strict(coord, text: str, clear_repeats: int = 7):
    click_abs(coord)
    time.sleep(FIELD_DELAY)
    send_keys("{DEL " + str(clear_repeats) + "}")
    send_keys("{BACKSPACE " + str(clear_repeats) + "}")
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


def get_current_month_value() -> str:
    now = datetime.now()
    return f"{now.month:02d}/{now.year}"


def mark_column(coord, sequence: str):
    for _ in range(SEQUENCE_REPEAT):
        context_menu_sequence(coord, sequence, delay_s=0.2, post_delay_s=0.5)




def sanitize_filename(name: str) -> str:
    invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid:
        name = name.replace(ch, "-")
    return " ".join(name.split()).strip()


def normalize_text(value) -> str:
    text = str(value or "")
    return " ".join(text.split()).strip()


def strip_procuracao(name: str) -> str:
    text = normalize_text(name)
    lowered = text.lower()
    if lowered.endswith(" - procuração") or lowered.endswith(" - procuracao"):
        return text[: text.rfind(" - ")].strip()
    return text

def normalize_text_lower(text: str) -> str:
    return (text or "").strip().lower()


def find_window_by_title(keywords):
    desk = Desktop(backend="uia")
    for w in desk.windows():
        title = normalize_text_lower(w.window_text())
        if not title:
            continue
        if any(k in title for k in keywords):
            return w
    return None


def find_button_by_keywords(root, keywords):
    buttons = root.descendants(control_type="Button")
    for b in buttons:
        name = normalize_text_lower(b.window_text() or b.element_info.name)
        if any(k in name for k in keywords):
            return b
    return None


def select_output_directory(path: str, coords: dict):
    click_abs(coords["btn_diretorio"])
    time.sleep(0.6)
    send_keys("%d")
    time.sleep(0.2)
    set_clipboard_text(path)
    send_keys("^a")
    send_keys("^v")
    time.sleep(0.2)
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




def resolve_excel_path(path: str) -> str:
    if path and os.path.exists(path):
        return path
    base = os.path.dirname(EXCEL_PREP_PATH)
    candidates = []
    if os.path.isdir(base):
        for name in os.listdir(base):
            if name.lower().endswith(".xlsx") and "preparado" in name.lower():
                candidates.append(os.path.join(base, name))
    if not candidates:
        return path
    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]


def load_company_map() -> dict:
    try:
        import pandas as pd
    except Exception as exc:
        raise SystemExit("pandas is required to read the Excel file") from exc
    excel_path = resolve_excel_path(EXCEL_PREP_PATH)
    if not excel_path or not os.path.exists(excel_path):
        raise SystemExit("Excel preparado nao encontrado.")
    try:
        df = pd.read_excel(excel_path, sheet_name=EXCEL_SHEET)
    except Exception:
        df = pd.read_excel(excel_path, sheet_name=0)
    if EXCEL_COL_NUM not in df.columns or EXCEL_COL_NOME not in df.columns:
        raise SystemExit("Colunas N e Nome nao encontradas no excel preparado.")
    mapping = {}
    for _, row in df.iterrows():
        num = str(row.get(EXCEL_COL_NUM, "")).strip()
        if not num or not num.isdigit():
            continue
        name = strip_procuracao(str(row.get(EXCEL_COL_NOME, "")).strip())
        if not name:
            continue
        mapping[num.zfill(4)] = name
    return mapping


def find_latest_period_dir(base_dir: str) -> str | None:
    if not os.path.isdir(base_dir):
        return None
    entries = [
        os.path.join(base_dir, d)
        for d in os.listdir(base_dir)
        if os.path.isdir(os.path.join(base_dir, d)) and d.isdigit()
    ]
    if not entries:
        return None
    entries.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return entries[0]


def find_folha_pdf(company_dir: str) -> str | None:
    hol_dir = os.path.join(company_dir, "FOLHAMENSAL", "HOLERITH")
    if not os.path.isdir(hol_dir):
        return None
    pdfs = [os.path.join(hol_dir, f) for f in os.listdir(hol_dir) if f.lower().endswith(".pdf")]
    if not pdfs:
        return None
    pdfs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return pdfs[0]


def unique_path(path: str) -> str:
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    idx = 2
    while True:
        candidate = f"{root} ({idx}){ext}"
        if not os.path.exists(candidate):
            return candidate
        idx += 1


def organize_outputs():
    period_dir = find_latest_period_dir(OUTPUT_DIR)
    if not period_dir:
        log("[organize] pasta de periodo nao encontrada.")
        return
    period = os.path.basename(period_dir)
    if len(period) != 6:
        log(f"[organize] periodo invalido: {period}")
        return
    year = period[:4]
    month = period[4:6]
    company_map = load_company_map()
    for num_dir in os.listdir(period_dir):
        src_company_dir = os.path.join(period_dir, num_dir)
        if not os.path.isdir(src_company_dir):
            continue
        name = company_map.get(num_dir)
        if not name:
            continue
        pdf_path = find_folha_pdf(src_company_dir)
        if not pdf_path:
            continue
        company = sanitize_filename(name)
        dest_dir = os.path.join(AUTOMATIZADO_ROOT, year, month, company)
        os.makedirs(dest_dir, exist_ok=True)
        dest_name = f"Folha de Pagamento - {company}.pdf"
        dest_path = unique_path(os.path.join(dest_dir, dest_name))
        try:
            shutil.copy2(pdf_path, dest_path)
            log(f"[organize] copiado: {dest_path}")
        except Exception as exc:
            log(f"[organize] falha ao copiar {pdf_path}: {exc}")

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
        elif key in DEFAULT_COORDS:
            out[key] = DEFAULT_COORDS[key]
        else:
            out[key] = None
    return out


def merge_prolabore_defaults(coords: dict) -> dict:
    if not os.path.exists(PROLABORE_COORDS_PATH):
        return coords
    base = load_coords(PROLABORE_COORDS_PATH)
    for key in coords:
        if coords[key] is None and key in base:
            value = base.get(key)
            if isinstance(value, dict) and "x" in value and "y" in value:
                coords[key] = (int(value["x"]), int(value["y"]))
            elif isinstance(value, (list, tuple)) and len(value) == 2:
                coords[key] = (int(value[0]), int(value[1]))
    return coords


def validate_coords(coords: dict):
    missing = [k for k, v in coords.items() if not v]
    if missing:
        raise SystemExit(f"Missing coordinates: {', '.join(missing)}")


def run_flow(main_win, coords: dict, skip_processar: bool = False):
    main_win.set_focus()
    time.sleep(0.25)

    send_keys("%m")
    time.sleep(0.2)
    send_keys("e")
    time.sleep(0.5)
    click_abs((1568, 310))
    time.sleep(0.4)

    replace_text_at(coords["field_processamento"], PROCESSAMENTO_VALUE)

    replace_text_at_strict(coords["field_mes_ano"], get_current_month_value())
    send_keys("{ENTER}")
    time.sleep(0.5)

    replace_text_at(coords["field_gerar_por"], GERAR_POR_VALUE)
    send_keys("{ENTER}")
    time.sleep(0.5)

    replace_text_at(coords["field_grupo"], GRUPO_VALUE)
    click_abs(coords["field_gerar_por"])
    time.sleep(GROUP_WAIT_S)

    mark_column(coords["ctx_calc_folha"], MARK_SEQUENCE)
    mark_column(coords["ctx_apuracao"], UNMARK_SEQUENCE)
    mark_column(coords["ctx_guias"], UNMARK_SEQUENCE)
    mark_column(coords["ctx_holerite"], MARK_SEQUENCE)
    mark_column(coords["ctx_rel_folha"], UNMARK_SEQUENCE)

    time.sleep(UI_CLICK_GAP_S)
    click_abs(coords["menu_parametros"])
    time.sleep(UI_CLICK_GAP_S)
    click_abs(coords["menu_relatorio_mensal"])
    time.sleep(UI_CLICK_GAP_S)

    select_output_directory(OUTPUT_DIR, coords)

    if not skip_processar:
        click_abs(coords["btn_processar"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--title-contains", default=ATTACH_TITLE_CONTAINS)
    ap.add_argument("--coords-file", default=COORDS_PATH)
    ap.add_argument("--capture-coords", action="store_true")
    ap.add_argument("--skip-ui", action="store_true")
    ap.add_argument("--skip-processar", action="store_true")
    ap.add_argument("--skip-organize", action="store_true")
    args = ap.parse_args()

    if args.capture_coords:
        guided_capture_coords(args.coords_file)
        return

    if args.skip_ui:
        return

    main_win = attach_main_window(args.title_contains)
    if not main_win:
        raise SystemExit("Main IOB window not found. Adjust --title-contains.")
    if not focus_main_window(args.title_contains):
        raise SystemExit("Unable to focus Folha de Pagamento window.")

    coords_data = load_coords(args.coords_file)
    coords = normalize_coords(coords_data)
    coords = merge_prolabore_defaults(coords)
    validate_coords(coords)
    run_flow(main_win, coords, skip_processar=args.skip_processar)
    if not args.skip_organize:
        organize_outputs()


if __name__ == "__main__":
    main()






