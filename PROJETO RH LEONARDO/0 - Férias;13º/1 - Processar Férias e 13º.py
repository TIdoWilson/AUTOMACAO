import argparse
import ctypes
import ctypes.wintypes as wintypes
import json
import os
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
PROCESSAMENTO_VALUE = "11"
GERAR_POR_VALUE = "1"
GRUPO_VALUE = "20"

COORDS_PATH = os.path.join(os.path.dirname(__file__), "coords_ferias13.json")

KEY_PAUSE = 0.04
CLICK_DELAY = 0.2
FIELD_DELAY = 0.2
UI_CLICK_GAP_S = 1.0

COLUMN_MENU_SEQUENCE = "DDRE"
COLUMN_MENU_REPEAT = 2

FINAL_OK_COORD = (1080, 600)
FINAL_CLOSE_COORD = (1471, 169)
FINAL_CLOSE_WAIT_S = 2.0
FINAL_BETWEEN_WAIT_S = 15.0
PROCESS_WAIT_S = 15.0


COORDS = {
    "field_processamento": None,
    "field_mes_ano": None,
    "field_gerar_por": None,
    "field_grupo": None,
    "col_calc_prov_ferias": None,
    "col_calc_prov_13": None,
    "btn_processar": None,
}

DEFAULT_COORDS = {
    "col_calc_prov_ferias": (917, 321),
    "col_calc_prov_13": (1010, 320),
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
    # Some fields ignore Ctrl+A; force clear with delete/backspace repeats.
    send_keys("{DEL " + str(clear_repeats) + "}")
    send_keys("{BACKSPACE " + str(clear_repeats) + "}")
    send_keys(text)
    time.sleep(FIELD_DELAY)


def context_menu_sequence(coord, moves: str, delay_s: float = 0.2, post_delay_s: float = 0.5):
    # Right-click (mouse button 2), then send the menu navigation sequence.
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


def get_previous_month_value() -> str:
    now = datetime.now()
    year = now.year
    month = now.month - 1
    if month == 0:
        month = 12
        year -= 1
    return f"{month:02d}/{year}"


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


def validate_coords(coords: dict):
    missing = [k for k, v in coords.items() if not v]
    if missing:
        raise SystemExit(f"Missing coordinates: {', '.join(missing)}")


def mark_column(coord):
    # Requires right-click (m2) and repeating the menu flow.
    for _ in range(COLUMN_MENU_REPEAT):
        context_menu_sequence(coord, COLUMN_MENU_SEQUENCE, delay_s=0.2, post_delay_s=0.5)


def run_iob_flow(main_win, coords: dict, skip_processar: bool = False, process_wait_s: float = PROCESS_WAIT_S):
    main_win.set_focus()
    time.sleep(0.25)

    send_keys("%m")
    time.sleep(0.2)
    send_keys("e")
    time.sleep(0.5)
    click_abs((1568, 310))
    time.sleep(0.4)

    replace_text_at(coords["field_processamento"], PROCESSAMENTO_VALUE)

    replace_text_at_strict(coords["field_mes_ano"], get_previous_month_value())
    send_keys("{ENTER}")
    time.sleep(0.5)

    replace_text_at(coords["field_gerar_por"], GERAR_POR_VALUE)
    send_keys("{ENTER}")
    time.sleep(0.5)

    replace_text_at(coords["field_grupo"], GRUPO_VALUE)
    click_abs(coords["field_gerar_por"])
    time.sleep(4.0)

    mark_column(coords["col_calc_prov_ferias"])
    mark_column(coords["col_calc_prov_13"])

    time.sleep(UI_CLICK_GAP_S)
    if not skip_processar:
        click_abs(coords["btn_processar"])
        time.sleep(process_wait_s)


def finalize_after_process():
    click_abs(FINAL_OK_COORD)
    time.sleep(FINAL_BETWEEN_WAIT_S)
    click_abs(FINAL_CLOSE_COORD)
    time.sleep(FINAL_CLOSE_WAIT_S)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--title-contains", default=ATTACH_TITLE_CONTAINS)
    ap.add_argument("--coords-file", default=COORDS_PATH)
    ap.add_argument("--capture-coords", action="store_true")
    ap.add_argument("--skip-ui", action="store_true")
    ap.add_argument("--skip-processar", action="store_true")
    ap.add_argument("--process-wait", type=float, default=PROCESS_WAIT_S)
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
    validate_coords(coords)
    run_iob_flow(
        main_win,
        coords,
        skip_processar=args.skip_processar,
        process_wait_s=args.process_wait,
    )
    if not args.skip_processar:
        finalize_after_process()


if __name__ == "__main__":
    main()
