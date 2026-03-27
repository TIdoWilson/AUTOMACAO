import json, os, time, math, ctypes, logging, argparse, threading
from ctypes import wintypes
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pyautogui
import uiautomation as auto

OUTPUT_FILENAME = "macro.json"
LOG_FILENAME = "record.log"

STOP_VK = 0x77  # F8

WAIT_THRESHOLD_MS = 250
DOUBLE_CLICK_MAX_MS = 350
DOUBLE_CLICK_MAX_DIST_PX = 4

MAX_ANCESTORS = 4
SEARCH_DEPTH_DEFAULT = 40
CHAIN_MAX_UP = 15

IGNORE_WINDOW_TITLE_CONTAINS = ["Accessibility Insights", "Inspect - Live"]

# --------------------- DPI awareness ---------------------
def set_dpi_awareness() -> str:
    try:
        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = ctypes.c_void_p(-4)
        ok = ctypes.windll.user32.SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
        if ok:
            return "PMv2"
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
        return "PM"
    except Exception:
        pass
    return "unknown"

# --------------------- COM (UIAutomation) ---------------------
_tls = threading.local()
COINIT_APARTMENTTHREADED = 0x2

def ensure_com_initialized(logger: logging.Logger):
    if getattr(_tls, "com_inited", False):
        return
    hr = ctypes.windll.ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)
    _tls.com_inited = True
    logger.debug("COM initialized (hr=%s)", hr)

# --------------------- Win32 APIs ---------------------
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# ---- Compat types ----
LRESULT = getattr(wintypes, "LRESULT", ctypes.c_ssize_t)
WPARAM  = getattr(wintypes, "WPARAM", ctypes.c_size_t)
LPARAM  = getattr(wintypes, "LPARAM", ctypes.c_ssize_t)
HMODULE = getattr(wintypes, "HMODULE", wintypes.HANDLE)
HHOOK   = wintypes.HANDLE
HKL     = getattr(wintypes, "HKL", wintypes.HANDLE)

class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("pt", POINT),
        ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p),
    ]

class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p),
    ]

# Hook constants
WH_KEYBOARD_LL = 13
WH_MOUSE_LL = 14
HC_ACTION = 0

WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105

WM_LBUTTONDOWN = 0x0201
WM_RBUTTONDOWN = 0x0204

GA_ROOT = 2

# ---- Win32 prototypes ----
user32.WindowFromPoint.argtypes = [POINT]
user32.WindowFromPoint.restype = wintypes.HWND

user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetAncestor.restype = wintypes.HWND

user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int

user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int

user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.GetWindowRect.restype = wintypes.BOOL

user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND

user32.GetKeyboardState.argtypes = [ctypes.POINTER(ctypes.c_ubyte)]
user32.GetKeyboardState.restype = wintypes.BOOL

user32.ToUnicodeEx.argtypes = [
    wintypes.UINT, wintypes.UINT,
    ctypes.POINTER(ctypes.c_ubyte),
    wintypes.LPWSTR, ctypes.c_int,
    wintypes.UINT, HKL
]
user32.ToUnicodeEx.restype = ctypes.c_int

user32.GetKeyboardLayout.argtypes = [wintypes.DWORD]
user32.GetKeyboardLayout.restype = HKL

user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD

# Module handle / errors
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
kernel32.GetModuleHandleW.restype = HMODULE

kernel32.GetLastError.argtypes = []
kernel32.GetLastError.restype = wintypes.DWORD

FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000
FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200
kernel32.FormatMessageW.argtypes = [
    wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
    wintypes.LPWSTR, wintypes.DWORD, ctypes.c_void_p
]
kernel32.FormatMessageW.restype = wintypes.DWORD

def win_errmsg(err: int) -> str:
    buf = ctypes.create_unicode_buffer(1024)
    flags = FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS
    n = kernel32.FormatMessageW(flags, None, err, 0, buf, len(buf), None)
    if n:
        return buf.value.strip()
    return ""

# ---- Hooks: IMPORTANT: SetWindowsHookEx must receive a function pointer type ----
HOOKPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, WPARAM, LPARAM)

user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, HMODULE, wintypes.DWORD]
user32.SetWindowsHookExW.restype = HHOOK

user32.UnhookWindowsHookEx.argtypes = [HHOOK]
user32.UnhookWindowsHookEx.restype = wintypes.BOOL

user32.CallNextHookEx.argtypes = [HHOOK, ctypes.c_int, WPARAM, LPARAM]
user32.CallNextHookEx.restype = LRESULT

user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.restype = wintypes.BOOL

user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.TranslateMessage.restype = wintypes.BOOL

user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.DispatchMessageW.restype = LRESULT

user32.PostQuitMessage.argtypes = [ctypes.c_int]
user32.PostQuitMessage.restype = None

# --------------------- window helpers ---------------------
def hwnd_from_point(x: int, y: int) -> int:
    return int(user32.WindowFromPoint(POINT(x, y)) or 0)

def root_hwnd(hwnd: int) -> int:
    return int(user32.GetAncestor(wintypes.HWND(hwnd), GA_ROOT) or 0)

def get_window_text(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(wintypes.HWND(hwnd), buf, 512)
    return buf.value

def get_class_name(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(wintypes.HWND(hwnd), buf, 256)
    return buf.value

def get_window_rect(hwnd: int) -> Optional[Tuple[int,int,int,int]]:
    r = wintypes.RECT()
    if not user32.GetWindowRect(wintypes.HWND(hwnd), ctypes.byref(r)):
        return None
    return (int(r.left), int(r.top), int(r.right), int(r.bottom))

def get_foreground_hwnd() -> int:
    h = user32.GetForegroundWindow()
    return int(h or 0)

def get_foreground_info() -> Dict[str, str]:
    hwnd = get_foreground_hwnd()
    if not hwnd:
        return {"name": "", "class_name": ""}
    return {"name": get_window_text(hwnd), "class_name": get_class_name(hwnd)}

def is_taskbar_window(class_name: str) -> bool:
    return class_name in ("Shell_TrayWnd", "MSTaskSwWClass", "MSTaskListWClass")

# --------------------- util ---------------------
def now_ms() -> int:
    return int(time.time() * 1000)

def clamp01(x: float) -> float:
    return 0.0 if x < 0.0 else (1.0 if x > 1.0 else x)

def dist(a: Tuple[int,int], b: Tuple[int,int]) -> float:
    return math.hypot(a[0]-b[0], a[1]-b[1])

def safe_str(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v)
    if not s or "Property does not exist" in s:
        return None
    return s

def normalize_control_type(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    s = str(s)
    return s[:-7] if s.endswith("Control") else s

def is_ignored_title(title: str) -> bool:
    low = (title or "").lower()
    return any(s.lower() in low for s in IGNORE_WINDOW_TITLE_CONTAINS)

def get_rect_uia(ctrl) -> Optional[Tuple[int,int,int,int]]:
    try:
        r = ctrl.BoundingRectangle
        return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        return None

def control_from_point(x: int, y: int, logger: logging.Logger):
    ensure_com_initialized(logger)
    if hasattr(auto, "ControlFromPoint"):
        return auto.ControlFromPoint(x, y)
    return auto.GetControlFromPoint(x, y)

def find_window_ancestor(ctrl) -> Optional[Any]:
    cur = ctrl
    for _ in range(80):
        if not cur:
            break
        try:
            if normalize_control_type(cur.ControlTypeName) == "Window":
                return cur
        except Exception:
            pass
        try:
            cur = cur.GetParentControl()
        except Exception:
            break
    return None

def is_generic_element(elem) -> bool:
    ctype = normalize_control_type(safe_str(getattr(elem, "ControlTypeName", None)))
    name  = safe_str(getattr(elem, "Name", None))
    cls   = safe_str(getattr(elem, "ClassName", None))
    if name:
        return False
    if ctype in ("Button", "Edit", "MenuItem", "Text", "ListItem", "TreeItem"):
        return False
    if cls and cls not in ("Window", "#32769"):
        return False
    return True

def pick_from_chain(base, win, logger: logging.Logger):
    cur = base
    for i in range(CHAIN_MAX_UP):
        if not cur:
            break
        if cur == win:
            break
        if not is_generic_element(cur) and get_rect_uia(cur):
            logger.debug("pick_from_chain picked at hop=%d", i)
            return cur
        try:
            cur = cur.GetParentControl()
        except Exception:
            break
    return None

def collect_ancestors(ctrl, win) -> List[Dict[str, Any]]:
    out = []
    cur = ctrl
    for _ in range(80):
        try:
            cur = cur.GetParentControl()
        except Exception:
            break
        if not cur or cur == win:
            break
        name = safe_str(getattr(cur, "Name", None))
        if not name or name == "Área de Trabalho 1":
            continue
        out.append({
            "control_type": normalize_control_type(safe_str(getattr(cur, "ControlTypeName", None))),
            "name": name,
            "class_name": safe_str(getattr(cur, "ClassName", None)),
        })
        if len(out) >= MAX_ANCESTORS:
            break
    return out

def build_locator(ctrl, win) -> Dict[str, Any]:
    return {
        "control_type": normalize_control_type(safe_str(getattr(ctrl, "ControlTypeName", None))),
        "name": safe_str(getattr(ctrl, "Name", None)),
        "class_name": safe_str(getattr(ctrl, "ClassName", None)),
        "search_depth": SEARCH_DEPTH_DEFAULT,
        "ancestors": collect_ancestors(ctrl, win) + [{
            "control_type": "Window",
            "name": safe_str(getattr(win, "Name", None)),
            "class_name": safe_str(getattr(win, "ClassName", None)),
        }],
    }

# --------------------- logger ---------------------
def setup_logger(log_path: str, level: str) -> logging.Logger:
    logger = logging.getLogger("macro_recorder")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s.%(msecs)03d [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
    sh = logging.StreamHandler(); sh.setFormatter(fmt); logger.addHandler(sh)
    fh = logging.FileHandler(log_path, encoding="utf-8"); fh.setFormatter(fmt); logger.addHandler(fh)
    return logger

# --------------------- recorder state ---------------------
@dataclass
class RecorderState:
    steps: List[Dict[str, Any]]
    text_buffer: List[str]
    modifiers: set
    last_action_ms: int
    last_click_ms: int
    last_click_pos: Tuple[int, int]

state = RecorderState([], [], set(), now_ms(), 0, (0,0))
logger: logging.Logger
dpi_mode: str = "unknown"

def append_step(step: Dict[str, Any]):
    step = dict(step)
    step["id"] = len(state.steps) + 1
    state.steps.append(step)

def start_action():
    t = now_ms()
    dt = t - state.last_action_ms
    if dt >= WAIT_THRESHOLD_MS:
        append_step({"type": "wait", "ms": int(dt)})
        logger.debug("Recorded wait=%dms", dt)
    state.last_action_ms = t

def flush_text_buffer():
    if not state.text_buffer:
        return
    txt = "".join(state.text_buffer)
    state.text_buffer.clear()
    if txt:
        start_action()
        append_step({"type": "type_text", "text": txt})
        logger.info("Recorded type_text len=%d", len(txt))

def record_focus_window(info: Optional[Dict[str, str]] = None):
    info = info or get_foreground_info()
    if is_ignored_title(info.get("name", "")):
        return
    start_action()
    append_step({"type": "focus_window", "window": info})
    logger.info("Recorded focus_window name=%r class=%r", info.get("name"), info.get("class_name"))

def record_click(x: int, y: int, button: str, is_double: bool):
    flush_text_buffer()

    top = root_hwnd(hwnd_from_point(x, y))
    top_name = get_window_text(top)
    top_cls = get_class_name(top)
    top_rect = get_window_rect(top)

    logger.info("Mouse click (%s) at (%d,%d) top_window name=%r class=%r double=%s",
                button, x, y, top_name, top_cls, is_double)

    if is_ignored_title(top_name):
        return

    if is_taskbar_window(top_cls):
        logger.info("Click on taskbar -> recording focus_window (foreground)")
        time.sleep(0.15)
        record_focus_window()
        return

    try:
        t0 = time.perf_counter()
        base = control_from_point(x, y, logger)
        win = find_window_ancestor(base)
        if not win:
            raise RuntimeError("UIA: no window ancestor")

        elem = pick_from_chain(base, win, logger)
        if not elem:
            raise RuntimeError("UIA: chain could not find stable element")

        rect = get_rect_uia(elem)
        if not rect:
            raise RuntimeError("UIA: picked element without rect")

        l,t,r,b = rect
        rx = clamp01((x - l) / max(1, r-l))
        ry = clamp01((y - t) / max(1, b-t))

        if not top_rect:
            raise RuntimeError("Win32: no window rect for window_point")
        wl, wt, wr, wb = top_rect
        wrx = clamp01((x - wl) / max(1, wr - wl))
        wry = clamp01((y - wt) / max(1, wb - wt))

        loc = build_locator(elem, win)
        start_action()
        append_step({
            "type": "double_click_uia" if is_double else "click_uia",
            "button": button,
            "timeout_ms": 8000,
            "target": {"window": {"name": top_name, "class_name": top_cls}, "locator": loc},
            "point": {"rx": rx, "ry": ry},
            "window_point": {"rx": wrx, "ry": wry},
        })

        dt = (time.perf_counter() - t0) * 1000.0
        logger.info("Recorded click_uia type=%r name=%r class=%r rect=%s took=%.1fms",
                    loc.get("control_type"), loc.get("name"), loc.get("class_name"), rect, dt)
        return

    except Exception as e:
        logger.warning("Falling back to click_win because: %r", e)

    if not top_rect:
        logger.error("click_win fallback failed: cannot get window rect")
        return
    wl, wt, wr, wb = top_rect
    rxw = clamp01((x - wl) / max(1, wr - wl))
    ryw = clamp01((y - wt) / max(1, wb - wt))
    start_action()
    append_step({
        "type": "double_click_win" if is_double else "click_win",
        "button": button,
        "timeout_ms": 8000,
        "window": {"name": top_name, "class_name": top_cls},
        "point": {"rx": rxw, "ry": ryw}
    })
    logger.info("Recorded click_win rect=(%d,%d,%d,%d) rx=%.4f ry=%.4f", wl, wt, wr, wb, rxw, ryw)

# --------------------- keyboard recording ---------------------
VK_SPECIAL = {
    0x0D: "ENTER",
    0x09: "TAB",
    0x1B: "ESC",
    0x08: "BACKSPACE",
    0x2E: "DELETE",
    0x20: "SPACE",
    0x25: "LEFT",
    0x26: "UP",
    0x27: "RIGHT",
    0x28: "DOWN",
}
for i in range(1, 13):
    VK_SPECIAL[0x70 + (i - 1)] = f"F{i}"

VK_CTRL = {0x11, 0xA2, 0xA3}
VK_SHIFT = {0x10, 0xA0, 0xA1}
VK_ALT = {0x12, 0xA4, 0xA5}

def vk_to_token(vk: int) -> Optional[str]:
    if vk in VK_SPECIAL:
        return VK_SPECIAL[vk]
    if 0x30 <= vk <= 0x39:
        return chr(vk)
    if 0x41 <= vk <= 0x5A:
        return chr(vk)
    return None

def get_foreground_layout() -> HKL:
    hwnd = user32.GetForegroundWindow()
    pid = wintypes.DWORD(0)
    tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid)) if hwnd else 0
    return user32.GetKeyboardLayout(tid)

def vk_to_text(vk: int, scan: int) -> Optional[str]:
    st = (ctypes.c_ubyte * 256)()
    if not user32.GetKeyboardState(st):
        return None
    layout = get_foreground_layout()
    buf = ctypes.create_unicode_buffer(16)
    res = user32.ToUnicodeEx(vk, scan, st, buf, len(buf), 0, layout)
    if res > 0:
        s = buf.value[:res]
        if s and all(ord(ch) >= 32 for ch in s):
            return s
    return None

def record_key(keys: List[str]):
    flush_text_buffer()
    start_action()
    append_step({"type": "key", "keys": keys})
    logger.info("Recorded key=%s", "+".join(keys))

def on_key_down(vk: int, scan: int):
    if vk == STOP_VK:
        flush_text_buffer()
        logger.info("Stop key pressed (F8).")
        user32.PostQuitMessage(0)
        return

    if vk in VK_CTRL:
        state.modifiers.add("CTRL"); return
    if vk in VK_ALT:
        state.modifiers.add("ALT"); return
    if vk in VK_SHIFT:
        state.modifiers.add("SHIFT"); return

    if vk in VK_SPECIAL:
        record_key([VK_SPECIAL[vk]])
        return

    if ("CTRL" in state.modifiers) or ("ALT" in state.modifiers):
        tok = vk_to_token(vk) or f"VK_{vk}"
        mods = sorted(list(state.modifiers))
        record_key(mods + [tok.upper() if len(tok) == 1 else tok])
        return

    # texto normal
    start_action()
    ch = vk_to_text(vk, scan)
    if ch:
        state.text_buffer.append(ch)

def on_key_up(vk: int):
    if vk in VK_CTRL:
        state.modifiers.discard("CTRL")
    if vk in VK_ALT:
        state.modifiers.discard("ALT")
    if vk in VK_SHIFT:
        state.modifiers.discard("SHIFT")

# --------------------- hooks ---------------------
_kb_hook: Optional[int] = None
_ms_hook: Optional[int] = None
_kb_proc_ref = None
_ms_proc_ref = None

def install_hooks(logger: logging.Logger):
    global _kb_hook, _ms_hook, _kb_proc_ref, _ms_proc_ref

    @HOOKPROC
    def kb_proc(nCode, wParam, lParam):
        try:
            if nCode == HC_ACTION:
                kbd = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                if wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                    on_key_down(int(kbd.vkCode), int(kbd.scanCode))
                elif wParam in (WM_KEYUP, WM_SYSKEYUP):
                    on_key_up(int(kbd.vkCode))
        except Exception as e:
            logger.exception("Keyboard hook error: %r", e)
        return user32.CallNextHookEx(_kb_hook, nCode, wParam, lParam)

    @HOOKPROC
    def ms_proc(nCode, wParam, lParam):
        try:
            if nCode == HC_ACTION and wParam in (WM_LBUTTONDOWN, WM_RBUTTONDOWN):
                ms = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
                x, y = int(ms.pt.x), int(ms.pt.y)

                t = now_ms()
                pos = (x, y)
                is_left = (wParam == WM_LBUTTONDOWN)
                is_double = (
                    is_left
                    and (t - state.last_click_ms) <= DOUBLE_CLICK_MAX_MS
                    and dist(pos, state.last_click_pos) <= DOUBLE_CLICK_MAX_DIST_PX
                )
                state.last_click_ms = t
                state.last_click_pos = pos

                ensure_com_initialized(logger)
                record_click(x, y, "left" if is_left else "right", is_double if is_left else False)
        except Exception as e:
            logger.exception("Mouse hook error: %r", e)
        return user32.CallNextHookEx(_ms_hook, nCode, wParam, lParam)

    _kb_proc_ref = kb_proc
    _ms_proc_ref = ms_proc

    hmod = kernel32.GetModuleHandleW(None)
    if not hmod:
        err = int(kernel32.GetLastError())
        raise OSError(f"GetModuleHandleW(None) falhou: err={err} {win_errmsg(err)}")

    _kb_hook = user32.SetWindowsHookExW(WH_KEYBOARD_LL, _kb_proc_ref, hmod, 0)
    if not _kb_hook:
        err = int(kernel32.GetLastError())
        raise OSError(f"SetWindowsHookExW(WH_KEYBOARD_LL) falhou: err={err} {win_errmsg(err)}")

    _ms_hook = user32.SetWindowsHookExW(WH_MOUSE_LL, _ms_proc_ref, hmod, 0)
    if not _ms_hook:
        err = int(kernel32.GetLastError())
        user32.UnhookWindowsHookEx(_kb_hook)
        _kb_hook = None
        raise OSError(f"SetWindowsHookExW(WH_MOUSE_LL) falhou: err={err} {win_errmsg(err)}")

    logger.info("Hooks instalados (keyboard+mouse). Pressione F8 para parar.")

def uninstall_hooks():
    global _kb_hook, _ms_hook
    if _kb_hook:
        user32.UnhookWindowsHookEx(_kb_hook)
        _kb_hook = None
    if _ms_hook:
        user32.UnhookWindowsHookEx(_ms_hook)
        _ms_hook = None

def save_macro(base_dir: str):
    out_path = os.path.join(base_dir, OUTPUT_FILENAME)
    payload = {
        "meta": {"created_at_ms": now_ms(), "version": 7, "dpi_mode": dpi_mode},
        "steps": state.steps
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    logger.info("Saved macro: %s (steps=%d)", out_path, len(state.steps))

def message_loop():
    msg = wintypes.MSG()
    while True:
        ret = user32.GetMessageW(ctypes.byref(msg), 0, 0, 0)
        if ret == 0:
            break
        if ret == -1:
            break
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--log-level", default="DEBUG", choices=["DEBUG","INFO","WARNING","ERROR"])
    args = ap.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    global logger, dpi_mode
    logger = setup_logger(os.path.join(base_dir, LOG_FILENAME), args.log_level)

    dpi_mode = set_dpi_awareness()

    logger.info("RECORDER started. Pressione F8 para parar.")
    logger.info("Log file: %s", os.path.join(base_dir, LOG_FILENAME))
    logger.info("DPI awareness: %s", dpi_mode)

    ensure_com_initialized(logger)
    install_hooks(logger)

    try:
        message_loop()
    finally:
        uninstall_hooks()
        flush_text_buffer()
        save_macro(base_dir)

if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    main()
