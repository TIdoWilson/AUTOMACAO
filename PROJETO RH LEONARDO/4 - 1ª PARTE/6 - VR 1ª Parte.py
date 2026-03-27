import argparse
import ctypes
import ctypes.wintypes as wintypes
import csv
import json
import os
import re
import time
import shutil
from datetime import datetime, timedelta
from pathlib import Path

try:
    from pywinauto import Desktop, mouse
    from pywinauto.keyboard import send_keys as send_keys_raw
except Exception as exc:
    raise SystemExit("pywinauto is required for UI automation") from exc


# =========================
# CONFIG
# =========================

AUTOMATIZADO_BASE = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\1ª PARTE")
VR_OUTPUT_DIR = Path(
    r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\4 - 1ª PARTE\VR"
)
COORDS_PATH = Path(__file__).with_suffix(".json")

KEY_PAUSE = 0.04
CLICK_DELAY = 0.2
FIELD_DELAY = 0.2
UI_CLICK_GAP_S = 1.0

COORDS = {
    "field_empresa": None,
    "field_data_sistema": None,
    "field_estabelecimento": None,
    "vr_open_1": None,
    "vr_open_2": None,
    "vr_open_3": None,
    "vr_ok_fornecimento": None,
    "vr_ja_processada": None,
    "vr_recibo_1": None,
    "vr_recibo_2": None,
    "vr_recibo_3": None,
    "field_classificacao_vr": None,
    "chk_emitir_2_vr": None,
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


def guided_capture_coords(path: Path):
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


def last_day_of_month(year: int, month: int) -> datetime:
    base = datetime(year, month, 1)
    next_month = base.replace(day=28) + timedelta(days=4)
    return next_month - timedelta(days=next_month.day)


def format_date(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y")


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


def set_estabelecimento(coords: dict, value: str):
    click_abs(coords["field_estabelecimento"])
    time.sleep(FIELD_DELAY)
    send_keys("^a")
    send_keys("{DEL 20}")
    send_keys("{BACKSPACE 20}")
    send_keys(value)
    send_keys("{TAB}")
    time.sleep(FIELD_DELAY)


def handle_estabelecimento_popup(timeout_s: float = 1.5) -> bool:
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        win = Desktop(backend="uia").window(title_re="Aviso do Sistema")
        if win.exists(0.2):
            try:
                if "Estabelecimento não encontrado." in (win.window_text() or ""):
                    pass
            except Exception:
                pass
            try:
                btn = win.child_window(title="OK", control_type="Button")
                if btn.exists(0.2):
                    btn.click_input()
                    return True
            except Exception:
                pass
        time.sleep(0.2)
    return False


def run_vr(
    coords: dict,
    company_num: str,
    out_dir: Path,
    mes_ano: str,
    estabelecimentos: list[int] | None,
    grupo: str,
    year: str,
    month: str,
    code_map: dict[str, str],
) -> list[Path]:
    set_empresa(coords, company_num)
    today = datetime.now()
    try:
        mes, ano = mes_ano.split("/")
        date_value = format_date(last_day_of_month(int(ano), int(mes)))
    except Exception:
        date_value = format_date(last_day_of_month(today.year, today.month))
    set_data_sistema(coords, date_value)

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

    results = []
    try:
        mes, ano = mes_ano.split("/")
        label_mes_ano = f"{mes} {ano}"
    except Exception:
        label_mes_ano = mes_ano.replace("/", " ")

    if not estabelecimentos:
        estabelecimentos = [1]

    should_set_estab = len(estabelecimentos) > 1
    for idx in estabelecimentos:
        if coords.get("field_estabelecimento") and should_set_estab:
            set_estabelecimento(coords, str(idx))
            if handle_estabelecimento_popup():
                continue

        click_abs(coords["chk_emitir_2_vr"])
        click_abs(coords["btn_imprimir_vr"])

        before = set(out_dir.glob("*.pdf"))
        click_abs(coords["vr_btn_salvar_pdf"])
        replace_text_at_strict(coords["vr_field_path"], str(out_dir))
        send_keys("{ENTER}")
        click_abs((975, 867))

        suffix = f"-{idx}" if idx >= 2 else ""
        base_name = f"{company_num}{suffix} VALE REFEICAO {label_mes_ano}"
        replace_text_at_select_all(coords["vr_field_filename"], base_name)
        if coords.get("vr_chk_no_open"):
            click_abs(coords["vr_chk_no_open"])
        click_abs(coords["vr_chk_pdfa"])
        click_abs(coords["vr_btn_ok_save"])
        time.sleep(0.5)
        send_keys("{ENTER}")

        new_pdf = wait_for_new_pdf(out_dir, before)
        saved = None
        if new_pdf and wait_file_stable(new_pdf):
            target = unique_path(out_dir / f"{sanitize_filename(base_name)}.pdf")
            try:
                new_pdf.rename(target)
                saved = target
            except Exception:
                pass
        if saved:
            copy_vr_to_automatizado(
                saved,
                company_num,
                grupo,
                year,
                month,
                code_map,
            )
            results.append(saved)
        if coords.get("vr_btn_back"):
            click_abs(coords["vr_btn_back"])

    return results


def extract_code_from_1a_parte_filename(name: str) -> str:
    m = re.search(r"\s-\s*(\d+)\s-\s*folha", name, re.IGNORECASE)
    if not m:
        return ""
    try:
        return str(int(m.group(1)))
    except Exception:
        return m.group(1)


def code_folder_map_1a_parte(year: str, month: str, grupo: str) -> dict[str, str]:
    base = AUTOMATIZADO_BASE / year / month / str(grupo)
    if not base.exists():
        return {}
    mapping = {}
    for pdf in base.rglob("*.pdf"):
        code = extract_code_from_1a_parte_filename(pdf.name)
        if not code:
            continue
        parent = pdf.parent
        if parent.is_dir():
            mapping.setdefault(code, parent.name)
    return mapping


def copy_vr_to_automatizado(
    src: Path | None,
    code: str,
    grupo: str,
    year: str,
    month: str,
    code_map: dict[str, str],
):
    if not src or not src.exists():
        return
    code_key = str(int(code)) if str(code).isdigit() else str(code)
    company = code_map.get(code_key) or code_key
    target_dir = AUTOMATIZADO_BASE / year / month / str(grupo) / company
    target_dir.mkdir(parents=True, exist_ok=True)
    target_path = unique_path(target_dir / src.name)
    try:
        shutil.copy2(src, target_path)
        print(f"[ok] copiado: {target_path}", flush=True)
    except Exception:
        pass


def list_codes_from_1a_parte(
    year: str,
    month: str,
    grupos: list[str],
    grupos_por_codigo: dict[str, str] | None = None,
) -> dict[str, list[str]]:
    found: dict[str, set[str]] = {}
    for grupo in grupos:
        base = AUTOMATIZADO_BASE / year / month / str(grupo)
        if not base.exists():
            continue
        for pdf in base.rglob("*.pdf"):
            code = extract_code_from_1a_parte_filename(pdf.name)
            if code:
                found.setdefault(str(grupo), set()).add(code)
    out: dict[str, list[str]] = {}
    for grupo, codes in found.items():
        out[grupo] = sorted(codes, key=lambda x: int(x))
    return out


def parse_grupos(value: str | None) -> list[str]:
    if not value:
        return ["6", "13", "14"]
    parts = re.split(r"[;, ]+", value.strip())
    grupos = []
    for part in parts:
        if not part:
            continue
        if part not in ("6", "13", "14"):
            raise SystemExit("Grupo invalido. Validos: 6, 13, 14")
        if part not in grupos:
            grupos.append(part)
    if not grupos:
        raise SystemExit("Nenhum grupo informado.")
    return grupos


def resolve_year_month(year: str | None, month: str | None) -> tuple[str, str]:
    now = datetime.now()
    year = year or str(now.year)
    month = month or f"{now.month:02d}"
    return year, month


def parse_estabelecimentos(value: str) -> list[int]:
    if not value:
        return []
    parts = re.split(r"[;, ]+", value.strip())
    out = []
    for part in parts:
        if not part:
            continue
        if not part.isdigit():
            continue
        num = int(part)
        if num not in out:
            out.append(num)
    return out


def load_empresas_csv(path: Path) -> tuple[dict[str, list[int]], dict[str, str]]:
    if not path.exists():
        raise SystemExit(f"Arquivo nao encontrado: {path}")
    data: dict[str, list[int]] = {}
    grupos_por_codigo: dict[str, str] = {}
    grupo_atual = ""
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        header = next(reader, None)
        for row in reader:
            if not row:
                continue
            code = (row[0] or "").strip()
            m_grupo = re.match(r"^\s*GRUPO\s*(\d+)\s*$", code, re.IGNORECASE)
            if m_grupo:
                grupo_atual = m_grupo.group(1)
                continue
            if not code or not code.isdigit():
                continue
            estabs = []
            if len(row) >= 2:
                estabs = parse_estabelecimentos(row[1])
            code_norm = str(int(code))
            data[code_norm] = estabs
            if grupo_atual:
                grupos_por_codigo[code_norm] = grupo_atual
    return data, grupos_por_codigo


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--coords-file", default=str(COORDS_PATH))
    ap.add_argument("--capture-coords", action="store_true")
    ap.add_argument("--title-contains", default="folha de pagamento")
    ap.add_argument("--grupos", default=None, help="Grupos para ler (ex: 13;14).")
    ap.add_argument("--year", default=None, help="Ano (YYYY).")
    ap.add_argument("--month", default=None, help="Mes (MM).")
    ap.add_argument("--pad", type=int, default=4, help="Zera a esquerda do codigo (padrao: 4).")
    ap.add_argument(
        "--empresas-csv",
        default=str(Path(__file__).resolve().parent / "empresas.csv"),
        help="CSV com codigo e estabelecimentos (padrao: empresas.csv na pasta do script).",
    )
    args = ap.parse_args()

    if args.capture_coords:
        guided_capture_coords(Path(args.coords_file))
        return

    coords_data = load_coords(Path(args.coords_file))
    coords = normalize_coords(coords_data)

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
    if not coords.get("field_estabelecimento"):
        print("[warn] field_estabelecimento nao informado; seguindo sem selecionar estabelecimento.")
    else:
        required_vr.append("field_estabelecimento")
    validate_coords(coords, required_vr)

    year, month = resolve_year_month(args.year, args.month)
    grupos = parse_grupos(args.grupos)
    estabelecimentos_por_codigo, grupos_por_codigo = load_empresas_csv(Path(args.empresas_csv))
    codes_by_group = list_codes_from_1a_parte(year, month, grupos, grupos_por_codigo=grupos_por_codigo)
    if not codes_by_group:
        raise SystemExit(f"Nenhum codigo encontrado em {AUTOMATIZADO_BASE}\\{year}\\{month}")
    allowed_codes = set(estabelecimentos_por_codigo.keys())

    if not focus_main_window(args.title_contains):
        raise SystemExit("Unable to focus Folha de Pagamento window.")

    mes_ano = f"{month}/{year}"
    for grupo in grupos:
        codes = [
            code
            for code in codes_by_group.get(str(grupo), [])
            if str(int(code)) in allowed_codes
        ]
        if not codes:
            continue
        code_map = code_folder_map_1a_parte(year, month, str(grupo))
        out_dir = VR_OUTPUT_DIR / str(grupo)
        out_dir.mkdir(parents=True, exist_ok=True)
        for code in codes:
            code_fmt = code.zfill(args.pad) if args.pad else code
            run_vr(
                coords,
                code_fmt,
                out_dir,
                mes_ano,
                estabelecimentos_por_codigo.get(str(int(code)), []),
                str(grupo),
                year,
                month,
                code_map,
            )

    print("OK")


if __name__ == "__main__":
    main()
