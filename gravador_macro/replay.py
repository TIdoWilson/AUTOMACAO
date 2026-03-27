import json, os, time, ctypes, logging, argparse
from ctypes import wintypes
from typing import Any, Dict, Optional, Tuple, List

import pyautogui
import uiautomation as auto

INPUT_FILENAME = "macro.json"
LOG_FILENAME = "replay.log"

# --------------------- Key mapping ---------------------
KEY_NAME_TO_PYAUTOGUI = {
    "ENTER": "enter",
    "TAB": "tab",
    "ESC": "esc",
    "BACKSPACE": "backspace",
    "DELETE": "delete",
    "SPACE": "space",
    "UP": "up",
    "DOWN": "down",
    "LEFT": "left",
    "RIGHT": "right",
}
for i in range(1, 13):
    KEY_NAME_TO_PYAUTOGUI[f"F{i}"] = f"f{i}"

MODS = {"CTRL": "ctrl", "ALT": "alt", "SHIFT": "shift"}

def run_key_step(step: Dict[str, Any], logger: logging.Logger):
    keys = step.get("keys") or []
    if not keys:
        logger.warning("  key step sem 'keys'")
        return

    norm: List[str] = []
    for k in keys:
        k = str(k).strip()
        if k in MODS:
            norm.append(MODS[k])
        elif k in KEY_NAME_TO_PYAUTOGUI:
            norm.append(KEY_NAME_TO_PYAUTOGUI[k])
        else:
            if len(k) == 1:
                norm.append(k.lower())
            else:
                norm.append(k.lower())

    logger.info("  key=%s -> %s", "+".join(keys), "+".join(norm))

    if len(norm) >= 2 and any(x in ("ctrl", "alt", "shift") for x in norm[:-1]):
        pyautogui.hotkey(*norm)
        return

    if len(norm) == 1:
        pyautogui.press(norm[0])
        return

    for k in norm:
        pyautogui.press(k)

# --------------------- COM + DPI ---------------------
COINIT_APARTMENTTHREADED = 0x2
ctypes.windll.ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)

def set_dpi_awareness() -> str:
    try:
        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = ctypes.c_void_p(-4)
        ok = ctypes.windll.user32.SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
        if ok:
            return "PMv2"
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return "PM"
    except Exception:
        pass
    return "unknown"

user32 = ctypes.windll.user32

WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

# Win32 prototypes
user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
user32.FindWindowW.restype = wintypes.HWND

user32.EnumWindows.argtypes = [WNDENUMPROC, wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL

user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int

user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int

user32.IsIconic.argtypes = [wintypes.HWND]
user32.IsIconic.restype = wintypes.BOOL

user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
user32.ShowWindow.restype = wintypes.BOOL

user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.SetForegroundWindow.restype = wintypes.BOOL

user32.BringWindowToTop.argtypes = [wintypes.HWND]
user32.BringWindowToTop.restype = wintypes.BOOL

user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND

user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.GetWindowRect.restype = wintypes.BOOL

SW_RESTORE = 9
SW_MAXIMIZE = 3

TASKBAR_CLASSES = {"Shell_TrayWnd", "MSTaskSwWClass", "MSTaskListWClass"}

# --------------------- logger ---------------------
def setup_logger(log_path: str, level: str) -> logging.Logger:
    logger = logging.getLogger("macro_replay")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s.%(msecs)03d [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
    sh = logging.StreamHandler(); sh.setFormatter(fmt); logger.addHandler(sh)
    fh = logging.FileHandler(log_path, encoding="utf-8"); fh.setFormatter(fmt); logger.addHandler(fh)
    return logger

def clamp01(x: float) -> float:
    return 0.0 if x < 0.0 else (1.0 if x > 1.0 else x)

def normalize_control_type(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    s = str(s)
    return s[:-7] if s.endswith("Control") else s

def get_text(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(wintypes.HWND(hwnd), buf, 512)
    return buf.value

def get_class(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(wintypes.HWND(hwnd), buf, 256)
    return buf.value

def get_rect_hwnd(hwnd: int) -> Optional[Tuple[int,int,int,int]]:
    r = wintypes.RECT()
    if not user32.GetWindowRect(wintypes.HWND(hwnd), ctypes.byref(r)):
        return None
    return (int(r.left), int(r.top), int(r.right), int(r.bottom))

def get_rect_uia(ctrl) -> Optional[Tuple[int,int,int,int]]:
    try:
        r = ctrl.BoundingRectangle
        return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        return None

def click_in_rect(rect: Tuple[int,int,int,int], rx: float, ry: float, button: str, double: bool):
    l,t,r,b = rect
    w = max(1, r-l); h = max(1, b-t)
    x = int(l + w * clamp01(float(rx)))
    y = int(t + h * clamp01(float(ry)))
    if double:
        pyautogui.doubleClick(x, y, button=button)
    else:
        pyautogui.click(x, y, button=button)
    return x, y

def is_taskbar_spec(win_spec: Dict[str, Any]) -> bool:
    return (win_spec.get("class_name") or "") in TASKBAR_CLASSES

def enum_windows() -> List[int]:
    out: List[int] = []
    @WNDENUMPROC
    def cb(hwnd, lparam):
        out.append(int(hwnd))
        return True
    user32.EnumWindows(cb, 0)
    return out

def find_window_fuzzy(win_spec: Dict[str, Any]) -> int:
    cls = win_spec.get("class_name")
    name = win_spec.get("name")

    # FindWindowW retorna NULL -> ctypes entrega None (porque HWND é c_void_p)
    # então não pode fazer int(None)
    if cls is not None or name is not None:
        h = user32.FindWindowW(cls, name)
        hwnd = int(h) if h else 0
        if hwnd:
            return hwnd

    want_cls = (cls or "").strip()
    want_name = (name or "").strip().lower()

    for h in enum_windows():
        got_cls = get_class(h)
        got_name = get_text(h)

        if want_cls and got_cls != want_cls:
            continue
        if want_name and want_name not in (got_name or "").lower():
            continue
        return h

    return 0

def uia_control_from_point(x: int, y: int):
    if hasattr(auto, "ControlFromPoint"):
        return auto.ControlFromPoint(x, y)
    return auto.GetControlFromPoint(x, y)

def get_uia_window_from_hwnd(hwnd: int, expected_cls: Optional[str], expected_name: Optional[str]):
    try:
        if hasattr(auto, "ControlFromHandle"):
            c = auto.ControlFromHandle(hwnd)
            cur = c
            for _ in range(80):
                if not cur:
                    break
                if normalize_control_type(getattr(cur, "ControlTypeName", None)) == "Window":
                    if expected_cls and getattr(cur, "ClassName", "") != expected_cls:
                        pass
                    else:
                        return cur
                try:
                    cur = cur.GetParentControl()
                except Exception:
                    break
    except Exception:
        pass

    try:
        w = auto.WindowControl(searchDepth=2, Name=expected_name or "", ClassName=expected_cls or "")
        if w.Exists(0.2, 0.05):
            return w
    except Exception:
        pass
    return None

def focus_window(win_spec: Dict[str, Any], logger: logging.Logger, timeout_ms: int = 8000, window_mode: str = "keep"):
    t0 = time.perf_counter()
    hwnd = find_window_fuzzy(win_spec)
    if not hwnd:
        raise TimeoutError(f"Janela não encontrada: {win_spec}")

    iconic = bool(user32.IsIconic(wintypes.HWND(hwnd)))
    fg = int(user32.GetForegroundWindow() or 0)
    already_fg = (fg == hwnd)

    if window_mode == "maximize":
        user32.ShowWindow(wintypes.HWND(hwnd), SW_MAXIMIZE)
    elif window_mode == "restore":
        user32.ShowWindow(wintypes.HWND(hwnd), SW_RESTORE)
    else:
        if iconic:
            user32.ShowWindow(wintypes.HWND(hwnd), SW_RESTORE)

    if not already_fg:
        user32.BringWindowToTop(wintypes.HWND(hwnd))
        user32.SetForegroundWindow(wintypes.HWND(hwnd))

    deadline = time.time() + (timeout_ms / 1000.0)
    rect = None
    while time.time() < deadline:
        rect = get_rect_hwnd(hwnd)
        if rect and (rect[3] - rect[1]) > 10 and not bool(user32.IsIconic(wintypes.HWND(hwnd))):
            break
        time.sleep(0.02)

    dt = (time.perf_counter() - t0) * 1000.0
    rect = rect or get_rect_hwnd(hwnd)
    logger.debug("focus_window ok in %.1fms hwnd=%d iconic=%s mode=%s rect=%s",
                 dt, hwnd, iconic, window_mode, rect)
    return hwnd, rect

def is_match(c, loc: Dict[str, Any]) -> bool:
    want_type = normalize_control_type(loc.get("control_type"))
    got_type = normalize_control_type(getattr(c, "ControlTypeName", None))
    if want_type and got_type != want_type:
        return False
    if loc.get("name") is not None and (getattr(c, "Name", "") or "") != loc["name"]:
        return False
    if loc.get("class_name") is not None and (getattr(c, "ClassName", "") or "") != loc["class_name"]:
        return False
    return True

def ancestors_ok(ctrl, ancestors) -> bool:
    for a in ancestors or []:
        if a.get("name") == "Área de Trabalho 1":
            continue
        cur = ctrl
        ok = False
        for _ in range(80):
            try:
                cur = cur.GetParentControl()
            except Exception:
                break
            if not cur:
                break
            if is_match(cur, a):
                ok = True
                break
        if not ok:
            return False
    return True

def is_locator_specific(locator: Dict[str, Any]) -> bool:
    name = locator.get("name")
    cls = (locator.get("class_name") or "")
    ctype = (locator.get("control_type") or "")
    if name:
        return True
    if cls and cls not in ("TPanel", "Window", "#32769"):
        return True
    if ctype in ("Button", "Edit", "MenuItem", "ListItem", "TreeItem"):
        return True
    return False

def iter_descendants(root, max_depth: int):
    try:
        first = root.GetFirstChildControl()
    except Exception:
        first = None

    stack: List[Tuple[Any, int]] = []
    if first:
        stack.append((first, 1))

    while stack:
        ctrl, depth = stack.pop()
        if not ctrl:
            continue

        yield ctrl, depth

        if depth < max_depth:
            try:
                child = ctrl.GetFirstChildControl()
            except Exception:
                child = None
            if child:
                stack.append((child, depth + 1))

        try:
            sib = ctrl.GetNextSiblingControl()
        except Exception:
            sib = None
        if sib:
            stack.append((sib, depth))

def resolve_by_locator_tree(uia_win, locator: Dict[str, Any], logger: logging.Logger, timeout_ms: int):
    t0 = time.perf_counter()
    max_depth = int(locator.get("search_depth") or 40)
    deadline = time.time() + (timeout_ms / 1000.0)

    while time.time() < deadline:
        for ctrl, depth in iter_descendants(uia_win, max_depth):
            try:
                if is_match(ctrl, locator) and ancestors_ok(ctrl, locator.get("ancestors", [])):
                    rect = get_rect_uia(ctrl)
                    logger.debug("locator_tree resolved: type=%r name=%r class=%r rect=%s depth=%d took=%.1fms",
                                 normalize_control_type(getattr(ctrl, "ControlTypeName", None)),
                                 getattr(ctrl, "Name", None),
                                 getattr(ctrl, "ClassName", None),
                                 rect, depth, (time.perf_counter()-t0)*1000.0)
                    return ctrl
            except Exception:
                continue
        time.sleep(0.05)

    logger.debug("locator_tree not found in %.1fms", (time.perf_counter()-t0)*1000.0)
    return None

def resolve_by_hittest(win_rect, window_point, locator, logger: logging.Logger):
    wl, wt, wr, wb = win_rect
    sx = int(wl + (wr-wl) * clamp01(float(window_point["rx"])))
    sy = int(wt + (wb-wt) * clamp01(float(window_point["ry"])))

    t0 = time.perf_counter()
    base = uia_control_from_point(sx, sy)
    logger.debug("hit-test at (%d,%d) base type=%r name=%r class=%r",
                 sx, sy,
                 normalize_control_type(getattr(base, "ControlTypeName", None)),
                 getattr(base, "Name", None),
                 getattr(base, "ClassName", None))

    cur = base
    for hop in range(30):
        if cur and is_match(cur, locator) and ancestors_ok(cur, locator.get("ancestors", [])):
            logger.debug("hit-test resolved at hop=%d took=%.1fms", hop, (time.perf_counter()-t0)*1000.0)
            return cur, (sx, sy)
        try:
            cur = cur.GetParentControl()
        except Exception:
            break

    logger.debug("hit-test not resolved took=%.1fms", (time.perf_counter()-t0)*1000.0)
    return None, (sx, sy)

def maybe_skip_taskbar_click(steps, idx, logger):
    step = steps[idx]
    if step.get("type") != "click_win":
        return False
    if not is_taskbar_spec(step.get("window", {})):
        return False
    logger.info("  taskbar click skipped")
    return True

def run_macro(path: str, wait_scale: float, wait_cap_ms: Optional[int], window_mode: str, logger: logging.Logger, typing_interval: float):
    macro = json.load(open(path, "r", encoding="utf-8"))
    steps = macro.get("steps", [])
    dpi = (macro.get("meta") or {}).get("dpi_mode")
    logger.info("Loaded macro steps=%d wait_scale=%.3f wait_cap_ms=%s (dpi=%s)", len(steps), wait_scale, wait_cap_ms, dpi)

    i = 0
    while i < len(steps):
        step = steps[i]
        t = step["type"]
        sid = step.get("id", i+1)
        logger.info("Step %d/%d (id=%s): %s", i+1, len(steps), sid, t)

        if t == "wait":
            ms = int(step["ms"])
            ms2 = int(ms * wait_scale)
            if wait_cap_ms is not None:
                ms2 = min(ms2, int(wait_cap_ms))
            logger.info("  wait recorded=%dms applied=%dms", ms, ms2)
            time.sleep(ms2/1000.0)
            i += 1
            continue

        if t == "focus_window":
            _, rect = focus_window(step["window"], logger, 8000, window_mode=window_mode)
            logger.info("  focused name=%r class=%r rect=%s", step["window"].get("name"), step["window"].get("class_name"), rect)
            i += 1
            continue

        if t in ("click_win", "double_click_win"):
            if t == "click_win" and maybe_skip_taskbar_click(steps, i, logger):
                i += 1
                continue
            _, rect = focus_window(step["window"], logger, int(step.get("timeout_ms", 8000)), window_mode=window_mode)
            pt = step["point"]
            is_double = (t == "double_click_win")
            x, y = click_in_rect(rect, pt["rx"], pt["ry"], step.get("button", "left"), is_double)
            logger.info("  %s at (%d,%d) rect=%s", t, x, y, rect)
            i += 1
            continue

        if t in ("click_uia", "double_click_uia"):
            timeout_ms = int(step.get("timeout_ms", 8000))
            win_spec = step["target"]["window"]
            hwnd, win_rect = focus_window(win_spec, logger, timeout_ms, window_mode=window_mode)
            if not win_rect:
                raise RuntimeError("Sem rect da janela (Win32)")

            uia_win = get_uia_window_from_hwnd(hwnd, win_spec.get("class_name"), win_spec.get("name"))
            locator = step["target"]["locator"]

            resolved = None
            if uia_win and is_locator_specific(locator):
                resolved = resolve_by_locator_tree(uia_win, locator, logger, timeout_ms)

            if not resolved:
                window_point = step.get("window_point") or step.get("point")
                resolved, (sx, sy) = resolve_by_hittest(win_rect, window_point, locator, logger)

            if not resolved:
                if t == "double_click_uia":
                    pyautogui.doubleClick(sx, sy, button=step.get("button", "left"))
                else:
                    pyautogui.click(sx, sy, button=step.get("button", "left"))
                logger.warning("  %s not resolved -> fallback screen click (%d,%d)", t, sx, sy)
                i += 1
                continue

            r = get_rect_uia(resolved) or win_rect
            pt = step["point"]
            is_double = (t == "double_click_uia")
            x, y = click_in_rect(r, pt["rx"], pt["ry"], step.get("button", "left"), is_double)
            logger.info("  %s resolved -> clicked (%d,%d) elem_rect=%s", t, x, y, r)
            i += 1
            continue

        if t == "type_text":
            text = step.get("text", "")
            logger.info("  type_text len=%d", len(text))
            pyautogui.write(text, interval=typing_interval)
            i += 1
            continue

        if t == "key":
            run_key_step(step, logger)
            i += 1
            continue

        raise ValueError(f"Step desconhecido: {t}")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--log-level", default="DEBUG", choices=["DEBUG","INFO","WARNING","ERROR"])
    ap.add_argument("--wait-scale", type=float, default=1.0)
    ap.add_argument("--wait-cap-ms", type=int, default=None)
    ap.add_argument("--window-mode", default="keep", choices=["keep","restore","maximize"])
    ap.add_argument("--typing-interval", type=float, default=0.0, help="intervalo entre caracteres no pyautogui.write")
    args = ap.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    logger = setup_logger(os.path.join(base_dir, LOG_FILENAME), args.log_level)

    dpi_mode = set_dpi_awareness()
    logger.info("DPI awareness: %s", dpi_mode)

    run_macro(os.path.join(base_dir, INPUT_FILENAME), args.wait_scale, args.wait_cap_ms, args.window_mode, logger, args.typing_interval)

if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    main()
