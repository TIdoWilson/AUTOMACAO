import argparse
import ctypes
import ctypes.wintypes as wintypes
import json
import os
import re
import unicodedata
import shutil
import time
from datetime import datetime, timedelta
from pathlib import Path

try:
    import pandas as pd
except Exception as exc:
    raise SystemExit("pandas is required to read the Excel file") from exc

try:
    from pywinauto import Desktop, mouse
    from pywinauto.keyboard import send_keys as send_keys_raw
except Exception as exc:
    raise SystemExit("pywinauto is required for UI automation") from exc


# =========================
# CONFIG
# =========================

EXCEL_PATH = (
    "W:\\DOCUMENTOS ESCRITORIO\\INSTALACAO SISTEMA\\python\\PROJETO RH LEONARDO\\"
    "4 - 1\u00aa PARTE\\domesticas_preparado.xlsx"
)
SHEET_NAME = "Dom\u00e9sticas"
COL_NUM = "N"
COL_NOME = "Nome"
COL_VT_VR = "VT/VR"

VR_OUTPUT_DIR = (
    "W:\\DOCUMENTOS ESCRITORIO\\INSTALACAO SISTEMA\\python\\PROJETO RH LEONARDO\\"
    "4 - 1\u00aa PARTE\\VT e VR\\VR"
)
VT_OUTPUT_DIR = (
    "W:\\DOCUMENTOS ESCRITORIO\\INSTALACAO SISTEMA\\python\\PROJETO RH LEONARDO\\"
    "4 - 1\u00aa PARTE\\VT e VR\\VT"
)

AUTOMATIZADO_BASE = "W:\\DOCUMENTOS ESCRITORIO\\RH\\AUTOMATIZADO\\DOM\u00c9STICAS"

COORDS_PATH = os.path.join(os.path.dirname(__file__), "coords_domesticas_beneficios.json")

KEY_PAUSE = 0.04
CLICK_DELAY = 0.2
FIELD_DELAY = 0.2
UI_CLICK_GAP_S = 1.0


COORDS = {
    "field_empresa": None,
    "field_data_sistema": None,
    "vr_open_1": None,
    "vr_open_2": None,
    "vr_open_3": None,
    "vr_ok_fornecimento": None,
    "vr_ja_processada": None,
    "vr_recibo_1": None,
    "vr_recibo_2": None,
    "vr_recibo_3": None,
    "vt_open_1": None,
    "vt_open_2": None,
    "vt_open_3": None,
    "vt_recibo_1": None,
    "vt_recibo_2": None,
    "vt_recibo_3": None,
    "field_classificacao_vr": None,
    "chk_emitir_2_vr": None,
    "field_classificacao_vt": None,
    "chk_emitir_2_vt": None,
    "btn_imprimir_vr": None,
    "vr_btn_salvar_pdf": None,
    "vr_field_path": None,
    "vr_field_filename": None,
    "vr_chk_no_open": None,
    "vr_chk_pdfa": None,
    "vr_btn_ok_save": None,
    "vr_btn_back": None,
}


# =========================
# UTIL
# =========================

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


def replace_text_at_strict(coord, text: str, clear_repeats: int = 10):
    click_abs(coord)
    time.sleep(FIELD_DELAY)
    send_keys("{DEL " + str(clear_repeats) + "}")
    send_keys("{BACKSPACE " + str(clear_repeats) + "}")
    send_keys(text)
    time.sleep(FIELD_DELAY)


def replace_text_at_select_all(coord, text: str):
    click_abs(coord)
    time.sleep(FIELD_DELAY)
    send_keys("^a")
    send_keys(text)
    time.sleep(FIELD_DELAY)


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
    print(f"[coord] saved: {path}", flush=True)


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


def validate_coords(coords: dict, required_keys: list[str]):
    missing = [k for k in required_keys if not coords.get(k)]
    if missing:
        raise SystemExit(f"Missing coordinates: {', '.join(missing)}")


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


def sanitize_filename(name: str) -> str:
    invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid:
        name = name.replace(ch, "-")
    return " ".join(name.split()).strip()


def normalize_key(value: str) -> str:
    raw = (value or "").strip()
    normalized = unicodedata.normalize("NFKD", raw)
    cleaned = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    return " ".join(cleaned.lower().split())


def strip_procuracao(name: str) -> str:
    text = (name or "").strip()
    lowered = text.lower()
    if lowered.endswith(" - procura\u00e7\u00e3o") or lowered.endswith(" - procuracao"):
        return text[: text.rfind(" - ")].strip()
    return text


def cell_has_keyword(value, keyword: str) -> bool:
    if value is None:
        return False
    text = str(value).strip().lower()
    return keyword.lower() in text


def last_day_of_month(dt: datetime) -> datetime:
    next_month = dt.replace(day=28) + timedelta(days=4)
    return next_month - timedelta(days=next_month.day)


def format_date(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y")


def get_last_day_current_month() -> str:
    return format_date(last_day_of_month(datetime.now()))


def get_last_day_current_month_dt() -> datetime:
    return last_day_of_month(datetime.now())


def last_business_day_next_month() -> str:
    now = datetime.now()
    month_ahead = (now.replace(day=1) + timedelta(days=32)).replace(day=1)
    last_day = last_day_of_month(month_ahead)
    while last_day.weekday() >= 5:
        last_day -= timedelta(days=1)
    return format_date(last_day)


def last_business_day_next_month_dt() -> datetime:
    now = datetime.now()
    month_ahead = (now.replace(day=1) + timedelta(days=32)).replace(day=1)
    last_day = last_day_of_month(month_ahead)
    while last_day.weekday() >= 5:
        last_day -= timedelta(days=1)
    return last_day


def copy_to_automatizado(
    src: Path | None,
    name: str,
    when: datetime,
    filename_override: str | None = None,
) -> Path | None:
    if not src or not src.exists():
        return None
    clean_name = strip_procuracao(name)
    year = f"{when.year:04d}"
    month = f"{when.month:02d}"
    base_dir = Path(AUTOMATIZADO_BASE) / year / month
    base_dir.mkdir(parents=True, exist_ok=True)
    desired = sanitize_filename(clean_name)
    desired_key = normalize_key(desired)
    target_dir = base_dir / desired
    if not target_dir.exists():
        for entry in base_dir.iterdir():
            if entry.is_dir() and normalize_key(entry.name) == desired_key:
                target_dir = entry
                break
    target_dir.mkdir(parents=True, exist_ok=True)
    target_path = target_dir / (filename_override or src.name)
    if target_path.exists():
        print(f"[skip] ja existe: {target_path}", flush=True)
        return target_path
    try:
        shutil.copy2(src, target_path)
    except Exception:
        return None
    print(f"[ok] copiado: {target_path}", flush=True)
    return target_path


def get_automatizado_company_dir(name: str, when: datetime) -> Path:
    clean_name = strip_procuracao(name)
    year = f"{when.year:04d}"
    month = f"{when.month:02d}"
    base_dir = Path(AUTOMATIZADO_BASE) / year / month
    desired = sanitize_filename(clean_name)
    desired_key = normalize_key(desired)
    target_dir = base_dir / desired
    if target_dir.exists():
        return target_dir
    if base_dir.exists():
        for entry in base_dir.iterdir():
            if entry.is_dir() and normalize_key(entry.name) == desired_key:
                return entry
    return target_dir


def already_saved_in_automatizado(name: str, when: datetime, filename: str) -> bool:
    target_dir = get_automatizado_company_dir(name, when)
    target_path = target_dir / filename
    return target_path.exists()


def set_empresa(coords: dict, company_num: str):
    click_abs(coords["field_empresa"])
    time.sleep(FIELD_DELAY)
    send_keys("^a")
    send_keys("{DEL 5}")
    send_keys("{BACKSPACE 5}")
    send_keys(company_num)
    send_keys("{ENTER 2}")
    time.sleep(FIELD_DELAY)


def set_data_sistema(coords: dict, date_value: str):
    click_abs(coords["field_data_sistema"])
    time.sleep(FIELD_DELAY)
    send_keys("^a")
    send_keys("{DEL 40}")
    send_keys("{BACKSPACE 40}")
    send_keys(date_value)
    send_keys("{ENTER 3}")
    time.sleep(FIELD_DELAY)


def wait_for_new_pdf(folder: Path, before_files: set[Path], timeout_s: float = 30.0) -> Path | None:
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        current = set(folder.glob("*.pdf"))
        new_files = [p for p in current if p not in before_files]
        if new_files:
            newest = max(new_files, key=lambda p: p.stat().st_mtime)
            return newest
        time.sleep(0.3)
    return None


def wait_file_stable(path: Path, timeout_s: float = 30.0, stable_for_s: float = 1.5) -> bool:
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
        time.sleep(0.2)
    return False


def dismiss_sem_dados_para_listar(timeout_s: float = 8.0) -> bool:
    def _norm(msg: str) -> str:
        raw = (msg or "").strip().lower()
        txt = unicodedata.normalize("NFKD", raw)
        return "".join(ch for ch in txt if not unicodedata.combining(ch))

    def _has_no_data_text(msg: str) -> bool:
        t = _norm(msg)
        return "nao ha dados para listar" in t or ("dados" in t and "listar" in t and "nao" in t)

    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            desk = Desktop(backend="uia")
            for win in desk.windows():
                try:
                    title = win.window_text() or ""
                except Exception:
                    title = ""
                hit = _has_no_data_text(title)
                if not hit:
                    try:
                        for child in win.descendants():
                            txt = child.window_text() or ""
                            if _has_no_data_text(txt):
                                hit = True
                                break
                    except Exception:
                        pass
                if not hit:
                    continue
                try:
                    win.set_focus()
                except Exception:
                    pass
                # Tenta clicar no botao OK explicitamente.
                try:
                    for btn in win.descendants(control_type="Button"):
                        btxt = (btn.window_text() or "").strip().lower()
                        if btxt in ("ok", "&ok"):
                            try:
                                btn.click_input()
                                time.sleep(0.15)
                                return True
                            except Exception:
                                pass
                except Exception:
                    pass
                # Fallback por teclado.
                send_keys("{ENTER}")
                time.sleep(0.15)
                send_keys("{ENTER}")
                return True
        except Exception:
            pass
        time.sleep(0.15)
    return False


def wait_for_processamento_concluido(timeout_s: float = 60.0) -> bool:
    def _norm(msg: str) -> str:
        raw = (msg or "").strip().lower()
        txt = unicodedata.normalize("NFKD", raw)
        return "".join(ch for ch in txt if not unicodedata.combining(ch))

    def _click_ok(win) -> bool:
        try:
            for btn in win.descendants(control_type="Button"):
                btxt = _norm(btn.window_text() or "")
                if btxt in ("ok", "&ok"):
                    try:
                        btn.click_input()
                        time.sleep(0.2)
                        return True
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            win.set_focus()
        except Exception:
            pass
        send_keys("{ENTER}")
        time.sleep(0.2)
        return True

    deadline = time.time() + timeout_s
    while time.time() < deadline:
        if dismiss_sem_dados_para_listar(timeout_s=0.25):
            return False
        try:
            desk = Desktop(backend="uia")
            for win in desk.windows():
                texts = []
                try:
                    texts.append(win.window_text() or "")
                except Exception:
                    pass
                try:
                    for child in win.descendants():
                        txt = child.window_text() or ""
                        if txt:
                            texts.append(txt)
                except Exception:
                    pass
                merged = _norm(" | ".join(texts))
                is_aviso = "aviso do sistema" in merged
                if "processamento concluido" in merged:
                    if is_aviso:
                        return _click_ok(win)
                    # Alguns layouts mostram apenas a mensagem sem titulo visivel.
                    return _click_ok(win)
        except Exception:
            pass
        time.sleep(0.2)
    return False


def wait_file_ready(path: Path, timeout_s: float = 30.0) -> bool:
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        if path.exists() and wait_file_stable(path, timeout_s=5.0, stable_for_s=1.2):
            return True
        time.sleep(0.25)
    return False


def confirm_arquivo_salvo(timeout_s: float = 6.0) -> bool:
    def _norm(msg: str) -> str:
        raw = (msg or "").strip().lower()
        txt = unicodedata.normalize("NFKD", raw)
        return "".join(ch for ch in txt if not unicodedata.combining(ch))

    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            desk = Desktop(backend="uia")
            for win in desk.windows():
                texts = []
                try:
                    texts.append(win.window_text() or "")
                except Exception:
                    pass
                try:
                    for child in win.descendants():
                        txt = child.window_text() or ""
                        if txt:
                            texts.append(txt)
                except Exception:
                    pass
                merged = _norm(" | ".join(texts))
                if ("arquivo" in merged and "salvo" in merged) or "salvo com sucesso" in merged:
                    try:
                        win.set_focus()
                    except Exception:
                        pass
                    send_keys("{ENTER}")
                    time.sleep(0.2)
                    return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    base = path.stem
    suffix = path.suffix
    idx = 2
    while True:
        candidate = path.with_name(f"{base} ({idx}){suffix}")
        if not candidate.exists():
            return candidate
        idx += 1


def remove_suffix_counter(path: Path, desired: Path) -> tuple[Path, bool]:
    if path == desired:
        return path, False
    if desired.exists():
        return path, False
    try:
        path.rename(desired)
        return desired, True
    except Exception:
        return path, False


def find_pdf_by_base(folder: Path, base_name: str) -> Path | None:
    exact = folder / f"{base_name}.pdf"
    if exact.exists():
        return exact
    pattern = re.compile(rf"^{re.escape(base_name)} \(\d+\)\.pdf$", re.IGNORECASE)
    matches = [p for p in folder.glob("*.pdf") if pattern.match(p.name)]
    if not matches:
        return None
    return max(matches, key=lambda p: p.stat().st_mtime)


def find_pdf_recursive_by_base(root: Path, base_name: str) -> Path | None:
    exact = list(root.rglob(f"{base_name}.pdf"))
    if exact:
        return max(exact, key=lambda p: p.stat().st_mtime)
    pattern = re.compile(rf"^{re.escape(base_name)} \(\d+\)\.pdf$", re.IGNORECASE)
    matches = [p for p in root.rglob("*.pdf") if pattern.match(p.name)]
    if not matches:
        return None
    return max(matches, key=lambda p: p.stat().st_mtime)


def run_vr(coords: dict, company_num: str, name: str, out_dir: Path) -> Path | None:
    set_empresa(coords, company_num)
    set_data_sistema(coords, get_last_day_current_month())
    click_abs(coords["vr_open_1"])
    click_abs(coords["vr_open_2"])
    click_abs(coords["vr_open_3"])
    if coords.get("vr_ok_fornecimento"):
        click_abs(coords["vr_ok_fornecimento"])
        time.sleep(2.0)
        if coords.get("vr_ja_processada"):
            click_abs(coords["vr_ja_processada"])
            send_keys("{ENTER}")
    click_abs(coords["vr_recibo_1"])
    click_abs(coords["vr_recibo_2"])
    click_abs(coords["vr_recibo_3"])
    replace_text_at(coords["field_classificacao_vr"], "1")
    click_abs(coords["chk_emitir_2_vr"])
    click_abs(coords["btn_imprimir_vr"])
    time.sleep(0.35)
    if dismiss_sem_dados_para_listar():
        print(f"[skip] sem dados para listar (VR): {name}", flush=True)
        if coords.get("vr_btn_back"):
            click_abs(coords["vr_btn_back"])
        return None
    before = set(out_dir.glob("*.pdf"))
    click_abs(coords["vr_btn_salvar_pdf"])
    replace_text_at_strict(coords["vr_field_path"], str(out_dir))
    send_keys("{ENTER}")
    click_abs((975, 867))
    clean_name = strip_procuracao(name)
    replace_text_at_select_all(coords["vr_field_filename"], f"VR - {clean_name}")
    if coords.get("vr_chk_no_open"):
        click_abs(coords["vr_chk_no_open"])
    click_abs(coords["vr_chk_pdfa"])
    click_abs(coords["vr_btn_ok_save"])
    time.sleep(0.5)
    send_keys("{ENTER}")
    new_pdf = wait_for_new_pdf(out_dir, before)
    saved = None
    if new_pdf and wait_file_stable(new_pdf):
        target = unique_path(out_dir / f"VR - {sanitize_filename(clean_name)}.pdf")
        try:
            new_pdf.rename(target)
            saved = target
        except Exception:
            pass
    if coords.get("vr_btn_back"):
        click_abs(coords["vr_btn_back"])
    return saved


def run_vt(coords: dict, company_num: str, name: str, out_dir: Path) -> Path | None:
    set_empresa(coords, company_num)
    set_data_sistema(coords, last_business_day_next_month())
    clean_name = sanitize_filename(strip_procuracao(name))
    target = unique_path(out_dir / f"VT - {clean_name}.pdf")

    # Novo fluxo VT:
    # ALT+M, V, F
    send_keys("%m")
    time.sleep(0.2)
    send_keys("v")
    time.sleep(0.2)
    send_keys("f")
    time.sleep(0.2)

    # ALT+O e aguarda "Processamento concluído"
    send_keys("%o")
    time.sleep(0.2)
    if not wait_for_processamento_concluido(timeout_s=90.0):
        if dismiss_sem_dados_para_listar(timeout_s=1.5):
            print(f"[skip] sem dados para listar (VT): {name}", flush=True)
            return None

    # ALT+M, V, E
    send_keys("%m")
    time.sleep(0.2)
    send_keys("v")
    time.sleep(0.2)
    send_keys("e")
    time.sleep(0.2)

    # 2 TAB, digita 1 (0.1s entre cada TAB)
    for _ in range(2):
        send_keys("{TAB}")
        time.sleep(0.1)
    send_keys("1")
    time.sleep(0.2)

    # 15 TAB, ESPACO (0.1s entre cada TAB)
    for _ in range(15):
        send_keys("{TAB}")
        time.sleep(0.1)
    send_keys("{SPACE}")
    time.sleep(0.2)

    # ALT+I
    send_keys("%i")
    time.sleep(0.2)
    if dismiss_sem_dados_para_listar(timeout_s=2.0):
        print(f"[skip] sem dados para listar (VT): {name}", flush=True)
        return None

    # espera e imprime
    time.sleep(10.0)
    click_abs((400, 88), pre_wait_s=0.2)
    time.sleep(0.8)

    if dismiss_sem_dados_para_listar(timeout_s=2.0):
        print(f"[skip] sem dados para listar (VT): {name}", flush=True)
        return None

    # Salva digitando caminho completo + nome + extensao
    send_keys(str(target))
    time.sleep(0.2)
    send_keys("{ENTER}")

    if wait_file_ready(target, timeout_s=35.0):
        confirm_arquivo_salvo(timeout_s=6.0)
        click_abs((468, 85), pre_wait_s=0.2)
        return target
    return None


def resolve_excel_path(path: str) -> str:
    if path and os.path.exists(path):
        return path
    candidates = list(Path(os.path.dirname(__file__)).glob("*preparado*.xlsx"))
    if not candidates:
        return path
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return str(candidates[0])


def load_excel(path: str, sheet: str) -> pd.DataFrame:
    path = resolve_excel_path(path)
    if not path or not os.path.exists(path):
        raise SystemExit("Excel path not found. Set EXCEL_PATH or use --excel.")
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except ValueError:
        df = pd.read_excel(path, sheet_name=0)
    return df


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--coords-file", default=COORDS_PATH)
    ap.add_argument("--capture-coords", action="store_true")
    ap.add_argument("--excel", default=EXCEL_PATH)
    ap.add_argument("--sheet", default=SHEET_NAME)
    ap.add_argument("--mode", choices=["vr", "vt", "both"], default="both")
    ap.add_argument("--organize-only", action="store_true")
    ap.add_argument("--title-contains", default="folha de pagamento")
    args = ap.parse_args()

    if args.capture_coords:
        guided_capture_coords(args.coords_file)
        return

    coords_data = load_coords(args.coords_file)
    coords = normalize_coords(coords_data)

    df = load_excel(args.excel, args.sheet)
    missing_cols = [c for c in (COL_NUM, COL_NOME, COL_VT_VR) if c and c not in df.columns]
    if missing_cols:
        raise SystemExit(f"Missing columns in Excel: {', '.join(missing_cols)}")

    vr_dir = Path(VR_OUTPUT_DIR) if VR_OUTPUT_DIR else None
    vt_dir = Path(VT_OUTPUT_DIR) if VT_OUTPUT_DIR else None
    if args.mode in ("vr", "both") and not vr_dir:
        raise SystemExit("VR_OUTPUT_DIR is empty.")
    if args.mode in ("vt", "both") and not vt_dir:
        raise SystemExit("VT_OUTPUT_DIR is empty.")

    if vr_dir:
        vr_dir.mkdir(parents=True, exist_ok=True)
    if vt_dir:
        vt_dir.mkdir(parents=True, exist_ok=True)

    if args.organize_only:
        vr_date = get_last_day_current_month_dt()
        vt_date = get_last_day_current_month_dt()
        for _, row in df.iterrows():
            num = str(row.get(COL_NUM, "")).strip()
            name = str(row.get(COL_NOME, "")).strip()
            if not num or not num.isdigit():
                continue
            if not name:
                continue
            vt_vr_cell = row.get(COL_VT_VR)
            vr_flag = cell_has_keyword(vt_vr_cell, "vr")
            vt_flag = cell_has_keyword(vt_vr_cell, "vt")

            clean_name = sanitize_filename(strip_procuracao(name))
            raw_name = sanitize_filename(name)
            if args.mode in ("vr", "both") and vr_flag and vr_dir:
                vr_filename = f"VR - {clean_name}.pdf"
                if already_saved_in_automatizado(name, vr_date, vr_filename):
                    print(f"[skip] ja salvo no automatizado (VR): {name}", flush=True)
                else:
                    vr_src = find_pdf_recursive_by_base(vr_dir, f"VR - {clean_name}")
                    if not vr_src and raw_name != clean_name:
                        vr_src = find_pdf_recursive_by_base(vr_dir, f"VR - {raw_name}")
                    if vr_src:
                        vr_clean = vr_src.parent / vr_filename
                        vr_src, renamed = remove_suffix_counter(vr_src, vr_clean)
                        if renamed:
                            print(f"[ok] renomeado: {vr_src}", flush=True)
                        copy_to_automatizado(vr_src, name, vr_date, vr_filename)

            if args.mode in ("vt", "both") and vt_flag and vt_dir:
                vt_date_label = last_business_day_next_month_dt().strftime("%d-%m-%Y")
                vt_filename = f"VT - {clean_name} - {vt_date_label}.pdf"
                if already_saved_in_automatizado(name, vt_date, vt_filename):
                    print(f"[skip] ja salvo no automatizado (VT): {name}", flush=True)
                else:
                    vt_src = find_pdf_recursive_by_base(vt_dir, f"VT - {clean_name}")
                    if not vt_src and raw_name != clean_name:
                        vt_src = find_pdf_recursive_by_base(vt_dir, f"VT - {raw_name}")
                    if vt_src:
                        vt_clean = vt_src.parent / f"VT - {clean_name}.pdf"
                        vt_src, renamed = remove_suffix_counter(vt_src, vt_clean)
                        if renamed:
                            print(f"[ok] renomeado: {vt_src}", flush=True)
                        copy_to_automatizado(
                            vt_src,
                            name,
                            vt_date,
                            vt_filename,
                        )
        print("OK")
        return

    required_vr = [
        "field_empresa",
        "field_data_sistema",
        "vr_open_1",
        "vr_open_2",
        "vr_open_3",
        "vr_recibo_1",
        "vr_recibo_2",
        "vr_recibo_3",
        "field_classificacao_vr",
        "chk_emitir_2_vr",
        "btn_imprimir_vr",
        "vr_btn_salvar_pdf",
        "vr_field_path",
        "vr_field_filename",
        "vr_chk_pdfa",
        "vr_btn_ok_save",
    ]
    required_vt = [
        "field_empresa",
        "field_data_sistema",
        "vt_open_1",
        "vt_open_2",
        "vt_open_3",
        "vt_recibo_1",
        "vt_recibo_2",
        "vt_recibo_3",
        "field_classificacao_vt",
        "chk_emitir_2_vt",
        "btn_imprimir_vr",
        "vr_btn_salvar_pdf",
        "vr_field_path",
        "vr_field_filename",
        "vr_chk_pdfa",
        "vr_btn_ok_save",
    ]
    if args.mode == "vr":
        validate_coords(coords, required_vr)
    elif args.mode == "vt":
        validate_coords(coords, required_vt)
    else:
        validate_coords(coords, required_vr + required_vt)

    if not focus_main_window(args.title_contains):
        raise SystemExit("Unable to focus Folha de Pagamento window.")

    for _, row in df.iterrows():
        num = str(row.get(COL_NUM, "")).strip()
        name = str(row.get(COL_NOME, "")).strip()
        if not num or not num.isdigit():
            continue
        if not name:
            continue
        vt_vr_cell = row.get(COL_VT_VR)
        vr_flag = cell_has_keyword(vt_vr_cell, "vr")
        vt_flag = cell_has_keyword(vt_vr_cell, "vt")

        vr_saved = None
        vt_saved = None

        if args.mode in ("vr", "both") and vr_flag:
            vr_date = get_last_day_current_month_dt()
            vr_base = f"VR - {sanitize_filename(strip_procuracao(name))}.pdf"
            if already_saved_in_automatizado(name, vr_date, vr_base):
                print(f"[skip] ja salvo no automatizado (VR): {name}", flush=True)
            else:
                vr_saved = run_vr(coords, num.zfill(4), name, vr_dir)
                if not vr_saved and vr_dir:
                    vr_saved = find_pdf_recursive_by_base(vr_dir, vr_base.replace(".pdf", ""))
                    if not vr_saved:
                        raw_base = f"VR - {sanitize_filename(name)}.pdf"
                        vr_saved = find_pdf_recursive_by_base(vr_dir, raw_base.replace(".pdf", ""))
                if vr_saved:
                    vr_saved, renamed = remove_suffix_counter(vr_saved, vr_saved.parent / vr_base)
                    if renamed:
                        print(f"[ok] renomeado: {vr_saved}", flush=True)
                copy_to_automatizado(vr_saved, name, vr_date, vr_base)

        if args.mode in ("vt", "both") and vt_flag:
            vt_date = get_last_day_current_month_dt()
            vt_date_label = last_business_day_next_month_dt().strftime("%d-%m-%Y")
            vt_filename = f"VT - {sanitize_filename(strip_procuracao(name))} - {vt_date_label}.pdf"
            if already_saved_in_automatizado(name, vt_date, vt_filename):
                print(f"[skip] ja salvo no automatizado (VT): {name}", flush=True)
            else:
                vt_saved = run_vt(coords, num.zfill(4), name, vt_dir)
                vt_base = f"VT - {sanitize_filename(strip_procuracao(name))}.pdf"
                if not vt_saved and vt_dir:
                    vt_saved = find_pdf_recursive_by_base(vt_dir, vt_base.replace(".pdf", ""))
                    if not vt_saved:
                        raw_base = f"VT - {sanitize_filename(name)}.pdf"
                        vt_saved = find_pdf_recursive_by_base(vt_dir, raw_base.replace(".pdf", ""))
                if vt_saved:
                    vt_saved, renamed = remove_suffix_counter(vt_saved, vt_saved.parent / vt_base)
                    if renamed:
                        print(f"[ok] renomeado: {vt_saved}", flush=True)
                copy_to_automatizado(vt_saved, name, vt_date, vt_filename)

    print("OK")


if __name__ == "__main__":
    main()
