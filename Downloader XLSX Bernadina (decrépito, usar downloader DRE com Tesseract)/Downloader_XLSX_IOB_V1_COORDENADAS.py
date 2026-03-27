# iob_exporter.py
# -*- coding: utf-8 -*-

import os
import time
import unicodedata
import argparse
import json
from typing import List, Optional, Iterable

import psutil
from pywinauto import Application, Desktop, mouse
import ctypes
import ctypes.wintypes as wintypes
from pywinauto.keyboard import send_keys as send_keys_raw
from pywinauto.base_wrapper import BaseWrapper


# =========================
# CONFIG (igual ao seu C#)
# =========================

SAVE_ROOT = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\Downloader XLSX Bernadina\Excels Limpos IOB"

EMPRESAS = [721, 720]
ESTABS_721 = [1, 2, 3, 5, 6, 7]
RANGES = [("01/2025", "11/2025"), ("11/2025", "11/2025")]
GRAU = "5"

KEY_PAUSE = 0.05
CLICK_DELAY = 0.12
SHORT_DELAY = 0.08
FIELD_DELAY = 0.12
FIELD_GAP_DELAY = 0.35
CLEAR_MAX_CHARS = 12
PRINT_WAIT_BEFORE_SAVE = 5.0

ATTACH_PID = 6976  # ajuste se necessário
ATTACH_TITLE_CONTAINS = "contabilidade"

# AutomationId
AID_EMPRESA_EDIT = "67186"
AID_DRE_MENU = "211"

AID_BTN_IMPRIMIR = "460510"
AID_BTN_SALVAR_REL_EXCEL = "110848"
AID_BTN_FECHAR_PREV = "116232"

# DRE (AutomationId)
AID_DRE_ESTAB_EDIT = "4130570"
AID_DRE_MES_INI_EDIT = "14091910"
AID_DRE_MES_FIM_EDIT = "2491838"
AID_DRE_GRAU_EDIT = "3409660"
AID_DRE_DATA_EMISSAO_EDIT = "1705748"
AID_DRE_RADIO_ESTAB = ""         # opcional: radio "Estabelecimento"

# Coordenadas absolutas (uso principal)
COORD_EMPRESA = None
COORD_ESTAB = (791, 319)
COORD_MES_INI = (794, 338)
COORD_MES_FIM = (865, 337)
COORD_GRAU = (777, 363)
COORD_DATA_EMISSAO = (800, 396)
COORD_BTN_IMPRIMIR = (1195, 315)
COORD_BTN_SALVAR_EXCEL = (436, 92)
COORD_CKB_CLASSIFICACAO = (971, 471)
COORD_CKB_TITULOS_SALDO_ZERO = (686, 488)
COORD_CKB_CNPJ = (688, 575)
COORD_FECHAR_IMPRESSAO = (496, 87)

# Safety: avoid closing the wrong window by coordinates unless explicitly enabled.
ALLOW_UNSAFE_COORD_CLOSE = False
# Safety: abort if focus leaves the target app; avoid risky fallbacks.
SAFE_ABORT_ON_FOCUS_LOSS = True
STRICT_EMPRESA_FIELD = True
ASSUME_DRE_OPEN = True
MANUAL_CLICK_MODE = False

# Save dialog (Windows common dialog)
AID_SAVE_FILENAME = "1001"
AID_SAVE_BUTTON = "1"

COORDS_PATH = os.path.join(os.path.dirname(__file__), "coords.json")
COORD_FIELDS = {
    "empresa": "COORD_EMPRESA",
    "estab": "COORD_ESTAB",
    "mes_ini": "COORD_MES_INI",
    "mes_fim": "COORD_MES_FIM",
    "grau": "COORD_GRAU",
    "data_emissao": "COORD_DATA_EMISSAO",
    "btn_imprimir": "COORD_BTN_IMPRIMIR",
    "btn_salvar_excel": "COORD_BTN_SALVAR_EXCEL",
    "ckb_classificacao": "COORD_CKB_CLASSIFICACAO",
    "ckb_saldo_zero": "COORD_CKB_TITULOS_SALDO_ZERO",
    "ckb_cnpj": "COORD_CKB_CNPJ",
    "fechar_impressao": "COORD_FECHAR_IMPRESSAO",
}


# =========================
# UTIL
# =========================

def normalize(s: str) -> str:
    if not s:
        return ""
    s = s.strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s


def safe(fn, fallback=None):
    try:
        return fn()
    except Exception:
        return fallback


def log(msg: str):
    try:
        print(msg, flush=True)
    except Exception:
        pass


EXPECTED_PID: Optional[int] = None
EXPECTED_MAIN_WIN: Optional[BaseWrapper] = None


def set_expected_context(pid: Optional[int], main_win: Optional[BaseWrapper] = None):
    global EXPECTED_PID, EXPECTED_MAIN_WIN
    EXPECTED_PID = pid
    EXPECTED_MAIN_WIN = main_win


def _get_foreground_handle() -> Optional[int]:
    try:
        return int(ctypes.windll.user32.GetForegroundWindow())
    except Exception:
        return None


def ensure_expected_active(
    context: str,
    main_win: Optional[BaseWrapper] = None,
    require_active: bool = True,
):
    if not SAFE_ABORT_ON_FOCUS_LOSS:
        return
    target_win = main_win or EXPECTED_MAIN_WIN
    if EXPECTED_PID and not psutil.pid_exists(EXPECTED_PID):
        raise RuntimeError(f"[guard] target process ended during {context}; aborting.")
    if target_win and safe(lambda: target_win.exists(), True) is False:
        raise RuntimeError(f"[guard] main window closed during {context}; aborting.")
    if not require_active:
        return
    if EXPECTED_PID:
        active = safe(lambda: Desktop(backend="uia").get_active(), None)
        if not active:
            if target_win:
                fg = _get_foreground_handle()
                th = safe(lambda: target_win.element_info.handle, None)
                if fg and th and int(fg) == int(th):
                    return
                safe(lambda: target_win.set_focus(), None)
                time.sleep(SHORT_DELAY)
                active = safe(lambda: Desktop(backend="uia").get_active(), None)
        if not active:
            raise RuntimeError(
                f"[guard] active window not found during {context}; bring the app to the foreground."
            )
        active_pid = safe(lambda: active.element_info.process_id, None)
        if active_pid != EXPECTED_PID:
            if target_win:
                safe(lambda: target_win.set_focus(), None)
                time.sleep(SHORT_DELAY)
                active = safe(lambda: Desktop(backend="uia").get_active(), None)
                active_pid = safe(lambda: active.element_info.process_id, None) if active else None
                if active_pid != EXPECTED_PID:
                    fg = _get_foreground_handle()
                    th = safe(lambda: target_win.element_info.handle, None)
                    if fg and th and int(fg) == int(th):
                        return
            if active_pid != EXPECTED_PID:
                raise RuntimeError(f"[guard] focus left target app during {context}; aborting.")

def prompt_user_click(label: str):
    print(f">>> Clique em '{label}' e pressione ENTER aqui para continuar.")
    input()


def type_into_focused(text: str):
    clear_focused_text()
    send_keys(str(text), with_spaces=True)
    time.sleep(FIELD_GAP_DELAY)

def _coerce_coord(value):
    if isinstance(value, dict):
        if ("dx" in value and "dy" in value) or ("x" in value and "y" in value):
            return value
        return None
    if not isinstance(value, (list, tuple)) or len(value) != 2:
        return None
    try:
        return (int(value[0]), int(value[1]))
    except Exception:
        return None


def load_coords_from_file(path: str):
    if not path or not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        log(f"[coords] falha ao ler {path}: {e}")
        return
    if not isinstance(data, dict):
        log(f"[coords] formato invalido em {path}")
        return
    for key, var_name in COORD_FIELDS.items():
        if key not in data:
            continue
        coord = _coerce_coord(data[key])
        if coord:
            globals()[var_name] = coord


def save_coords_to_file(path: str, data: dict):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"[coords] falha ao salvar {path}: {e}")


def resolve_coord(coord, anchor_win: Optional[BaseWrapper] = None, require_inside: bool = False, context: str = "") -> Optional[tuple]:
    if coord is None:
        return None
    x = y = None
    if isinstance(coord, dict):
        if anchor_win and "dx" in coord and "dy" in coord:
            r = rect_of(anchor_win)
            if not r:
                return None
            x = r.left + int(coord["dx"])
            y = r.top + int(coord["dy"])
        elif "x" in coord and "y" in coord:
            x = int(coord["x"])
            y = int(coord["y"])
    elif isinstance(coord, (list, tuple)) and len(coord) == 2:
        x = int(coord[0])
        y = int(coord[1])
    if x is None or y is None:
        return None
    if anchor_win and require_inside:
        r = rect_of(anchor_win)
        if r and not (r.left <= x <= r.right and r.top <= y <= r.bottom):
            log(f"[coord] {context} fora da janela; abortando clique.")
            return None
    return (x, y)


def find_container_by_point(pid: Optional[int], x: int, y: int) -> Optional[BaseWrapper]:
    if not pid:
        return None
    candidates = []
    for w in iter_process_windows_uia(pid):
        search = [w]
        try:
            search += safe(lambda: w.descendants(control_type="Window"), []) or []
            search += safe(lambda: w.descendants(control_type="Pane"), []) or []
        except Exception:
            pass
        for sw in search:
            r = rect_of(sw)
            if not r:
                continue
            if r.left <= x <= r.right and r.top <= y <= r.bottom:
                area = max(1, (r.right - r.left) * (r.bottom - r.top))
                candidates.append((area, sw))
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def infer_anchor_from_coord(coord):
    if not isinstance(coord, dict):
        return None
    if "x" not in coord or "y" not in coord:
        return None
    return find_container_by_point(EXPECTED_PID, int(coord["x"]), int(coord["y"]))


def pick_anchor(
    coord,
    main_win: Optional[BaseWrapper] = None,
    dre_win: Optional[BaseWrapper] = None,
    report_win: Optional[BaseWrapper] = None,
):
    if isinstance(coord, dict):
        anchor = coord.get("anchor")
        if anchor == "main":
            return main_win or infer_anchor_from_coord(coord)
        if anchor == "dre":
            return dre_win or infer_anchor_from_coord(coord) or main_win
        if anchor == "report":
            if report_win:
                return report_win
            return dre_win or infer_anchor_from_coord(coord) or main_win
    return dre_win or report_win or main_win


def click_coord(coord, anchor_win: Optional[BaseWrapper] = None, context: str = ""):
    if not anchor_win:
        anchor_win = infer_anchor_from_coord(coord) or EXPECTED_MAIN_WIN
    abs_pos = resolve_coord(coord, anchor_win=None, require_inside=False, context=context)
    if anchor_win:
        r = rect_of(anchor_win)
        dx = dy = None
        if r and abs_pos:
            dx = int(abs_pos[0]) - r.left
            dy = int(abs_pos[1]) - r.top
        elif isinstance(coord, dict) and "dx" in coord and "dy" in coord:
            dx = int(coord["dx"])
            dy = int(coord["dy"])
        if r and dx is not None and dy is not None:
            width = r.right - r.left
            height = r.bottom - r.top
            if dx < 0 or dy < 0 or dx > width or dy > height:
                if abs_pos and r and (r.left <= abs_pos[0] <= r.right and r.top <= abs_pos[1] <= r.bottom):
                    dx = int(abs_pos[0]) - r.left
                    dy = int(abs_pos[1]) - r.top
                else:
                    if abs_pos and EXPECTED_MAIN_WIN:
                        mr = rect_of(EXPECTED_MAIN_WIN)
                        if mr and (mr.left <= abs_pos[0] <= mr.right and mr.top <= abs_pos[1] <= mr.bottom):
                            click_abs(abs_pos, require_active=False)
                            return
                    raise RuntimeError(f"Coordenada fora da janela: {context}")
            ensure_expected_active(f"click_coord:{context}", anchor_win, require_active=False)
            safe(lambda: anchor_win.set_focus(), None)
            time.sleep(SHORT_DELAY)
            try:
                anchor_win.click_input(coords=(dx, dy))
                time.sleep(CLICK_DELAY)
                return
            except Exception:
                pass
    if abs_pos and EXPECTED_MAIN_WIN:
        mr = rect_of(EXPECTED_MAIN_WIN)
        if mr and (mr.left <= abs_pos[0] <= mr.right and mr.top <= abs_pos[1] <= mr.bottom):
            click_abs(abs_pos, require_active=False)
            return
    pos = resolve_coord(coord, anchor_win, require_inside=bool(anchor_win), context=context)
    if not pos:
        raise RuntimeError(f"Coordenada invalida ou fora da janela: {context}")
    click_abs(pos, require_active=False)

def send_keys(keys: str, pause: Optional[float] = None, **kwargs):
    if pause is None:
        pause = KEY_PAUSE
    return send_keys_raw(keys, pause=pause, **kwargs)


def sanitize_filename(name: str) -> str:
    # Substitui caracteres inválidos em nomes de arquivo no Windows.
    invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid:
        name = name.replace(ch, "-")
    return name


def capture_click_coordinates(label: str, timeout_s: float = 30.0) -> Optional[tuple]:
    print(f"\n>>> Pressione ENTER e depois clique no campo '{label}' para capturar as coordenadas.")
    input("Pressione ENTER para armar a captura...")
    try:
        # aguarda o botao do mouse estar solto
        while ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000:
            time.sleep(0.01)
        t0 = time.time()
        while (time.time() - t0) < timeout_s:
            if ctypes.windll.user32.GetAsyncKeyState(0x01) & 0x8000:
                pt = wintypes.POINT()
                if ctypes.windll.user32.GetCursorPos(ctypes.byref(pt)):
                    print(f"[coord] {label}: x={pt.x} y={pt.y}")
                    return (pt.x, pt.y)
                raise OSError("GetCursorPos falhou")
            time.sleep(0.01)
        print(f"[coord] Timeout ao capturar coordenadas para {label}")
    except Exception as e:
        print(f"[coord] Falha ao capturar coordenadas para {label}: {e}")
    return None


def click_abs(pos, require_active: bool = True):
    ensure_expected_active("click_abs", require_active=require_active)
    try:
        mouse.click(button="left", coords=pos)
        time.sleep(CLICK_DELAY)
    except Exception:
        pass


def elem_text(el: BaseWrapper) -> str:
    return normalize(safe(lambda: el.window_text(), "") or "") or normalize(safe(lambda: el.element_info.name, "") or "")


def element_at_point(x: int, y: int) -> Optional[BaseWrapper]:
    try:
        return Desktop(backend="uia").from_point(x, y)
    except Exception:
        return None


def elem_has_keywords(el: BaseWrapper, keywords: List[str], max_depth: int = 6) -> bool:
    targets = [normalize(k) for k in keywords if k]
    cur = el
    depth = 0
    while cur is not None and depth < max_depth:
        txt = elem_text(cur)
        if any(t in txt for t in targets):
            return True
        cur = safe(lambda: cur.parent(), None)
        depth += 1
    return False


def elem_has_control_type(el: BaseWrapper, control_type: str, max_depth: int = 6) -> bool:
    target = normalize(control_type)
    cur = el
    depth = 0
    while cur is not None and depth < max_depth:
        ct = normalize(safe(lambda: cur.element_info.control_type, "") or "")
        if ct == target:
            return True
        cur = safe(lambda: cur.parent(), None)
        depth += 1
    return False


def find_container_by_offset(pid: Optional[int], coord, keywords: Optional[List[str]] = None) -> Optional[BaseWrapper]:
    if not pid or not isinstance(coord, dict):
        return None
    if "dx" not in coord or "dy" not in coord:
        return None
    try:
        dx = int(coord["dx"])
        dy = int(coord["dy"])
    except Exception:
        return None
    targets = keywords or []
    best = None
    for w in iter_process_windows_uia(pid):
        search = [w]
        try:
            search += safe(lambda: w.descendants(control_type="Window"), []) or []
            search += safe(lambda: w.descendants(control_type="Pane"), []) or []
        except Exception:
            pass
        for sw in search:
            r = rect_of(sw)
            if not r:
                continue
            x = r.left + dx
            y = r.top + dy
            if not (r.left <= x <= r.right and r.top <= y <= r.bottom):
                continue
            el = element_at_point(x, y)
            if not el:
                continue
            er = rect_of(el)
            if er and not (r.left <= er.left <= r.right and r.top <= er.top <= r.bottom):
                continue
            score = 0
            if targets and elem_has_keywords(el, targets, max_depth=6):
                score += 2
            if elem_has_control_type(el, "Button", max_depth=4):
                score += 1
            if score == 0:
                continue
            area = max(1, (r.right - r.left) * (r.bottom - r.top))
            cand = (-score, area, sw)
            if best is None or cand < best:
                best = cand
    return best[2] if best else None


def find_button_near_coord(
    root: BaseWrapper,
    coord_abs: tuple,
    keywords: List[str],
    max_dist: int = 120,
) -> Optional[BaseWrapper]:
    if not root or not coord_abs:
        return None
    x, y = coord_abs
    targets = [normalize(k) for k in keywords if k]
    best = None
    best_dist = None
    buttons = find_descendants(root, control_type="Button")
    for b in buttons:
        txt = elem_text(b)
        if not any(t in txt for t in targets):
            continue
        r = rect_of(b)
        if not r:
            continue
        cx = (r.left + r.right) // 2
        cy = (r.top + r.bottom) // 2
        dist = abs(cx - x) + abs(cy - y)
        if best_dist is None or dist < best_dist:
            best = b
            best_dist = dist
    if best is not None and best_dist is not None and best_dist <= max_dist:
        return best
    return None


def is_probable_save_target_at_point(pos: tuple) -> bool:
    el = element_at_point(pos[0], pos[1])
    if not el:
        return False
    if elem_has_keywords(el, ["salvar", "excel"], max_depth=6):
        return True
    if elem_has_keywords(el, ["relatorio", "impressao"], max_depth=6) and elem_has_control_type(el, "Button", max_depth=4):
        return True
    return False


def is_checked(cb: BaseWrapper) -> Optional[bool]:
    state = safe(lambda: cb.get_toggle_state(), None)
    if state is None:
        # fallback para LegacyAccessible pattern
        try:
            return cb.get_toggle_state() == 1
        except Exception:
            pass
    if state is not None:
        return state == 1
    return None


def set_checkbox_by_label(root: BaseWrapper, label_contains: List[str], should_check: bool) -> bool:
    targets = [normalize(x) for x in label_contains if x]
    checkboxes = find_descendants(root, control_type="CheckBox")
    for cb in checkboxes:
        txt = normalize(safe(lambda: cb.window_text(), "") or "") or normalize(safe(lambda: cb.element_info.name, "") or "")
        if any(t in txt for t in targets):
            desired = should_check
            curr = is_checked(cb)
            if curr is None or curr != desired:
                click_wrapper(cb)
                time.sleep(0.12)
            return True
    return False


def rect_of(w: BaseWrapper):
    r = safe(lambda: w.rectangle(), None)
    return r


def wait_until(timeout_s: float, step_s: float, predicate):
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        if predicate():
            return True
        time.sleep(step_s)
    return False


def iter_process_windows_uia(pid: int) -> List[BaseWrapper]:
    # Desktop(backend="uia").windows(process=pid) é o equivalente prático a "desktop.FindAllChildren + ProcessId"
    desk = Desktop(backend="uia")
    wins = safe(lambda: desk.windows(process=pid), []) or []
    return wins


def find_descendant_by_automation_id(root: BaseWrapper, automation_id: str) -> Optional[BaseWrapper]:
    if not automation_id:
        return None
    # pywinauto/UIA: automation_id fica em element_info.automation_id
    candidates = safe(lambda: root.descendants(), []) or []
    for c in candidates:
        aid = safe(lambda: c.element_info.automation_id, "") or ""
        if str(aid).strip().lower() == str(automation_id).strip().lower():
            return c
    return None


def find_descendants(root: BaseWrapper, control_type: Optional[str] = None) -> List[BaseWrapper]:
    if control_type:
        return safe(lambda: root.descendants(control_type=control_type), []) or []
    return safe(lambda: root.descendants(), []) or []


def click_wrapper(w: BaseWrapper):
    safe(lambda: w.set_focus(), None)
    time.sleep(CLICK_DELAY)
    safe(lambda: w.click_input(), None)
    time.sleep(CLICK_DELAY)


def click_at_offset(root: BaseWrapper, dx: int, dy: int):
    ensure_expected_active("click_at_offset")
    r = rect_of(root)
    if not r:
        return
    x = r.left + dx
    y = r.top + dy
    safe(lambda: root.click_input(coords=(x, y)), None)
    time.sleep(CLICK_DELAY)


def clear_focused_text(max_chars: int = CLEAR_MAX_CHARS):
    # Apaga de fato o conteudo antes de digitar algo novo (evita sobras).
    for _ in range(2):
        send_keys("^a{DEL}")
        time.sleep(SHORT_DELAY)
        # Alguns campos ignoram Ctrl+A; reforçamos com BACKSPACE no fim.
        send_keys(f"{{END}}{{BACKSPACE {max_chars}}}")
        time.sleep(SHORT_DELAY)


def select_all_fast(control: Optional[BaseWrapper], allow_foreign: bool = False):
    if control:
        safe(lambda: control.set_focus(), None)
        time.sleep(SHORT_DELAY)
        safe(lambda: control.double_click_input(), None)
    else:
        if not allow_foreign:
            try:
                mouse.double_click(button="left")
            except Exception:
                pass
    time.sleep(SHORT_DELAY)
    if not allow_foreign:
        ensure_expected_active("select_all_fast")
    send_keys("^a")
    time.sleep(SHORT_DELAY)


def replace_text(control: Optional[BaseWrapper], text: str, allow_foreign: bool = False):
    """
    Limpa o campo de forma agressiva e digita o novo valor.
    Usa os metodos do proprio controle quando possivel para garantir que o texto antigo seja removido.
    """
    def _normalized_current_value(ctrl: BaseWrapper) -> str:
        # Tenta extrair o valor atual do controle (quando suportado).
        current = safe(lambda: ctrl.get_value(), None)
        if current is None:
            current = safe(lambda: ctrl.window_text(), None)
        if current is None:
            current = safe(lambda: ctrl.element_info.name, None)
        return "" if current is None else str(current)

    def overwrite_fast(ctrl: Optional[BaseWrapper]):
        select_all_fast(ctrl, allow_foreign=allow_foreign)
        clear_focused_text()
        send_keys(str(text), with_spaces=True)
        time.sleep(FIELD_DELAY)

    if control:
        try:
            click_wrapper(control)
            # Tenta limpar via padroes do UIA (ValuePattern/EditPattern)
            target = "" if text is None else str(text)
            current = _normalized_current_value(control)
            if current == target:
                return
            try:
                control.set_edit_text(target)
                return
            except Exception:
                pass
            try:
                control.set_value(target)
                return
            except Exception:
                pass
            overwrite_fast(control)
            return
        except Exception:
            pass

    # Fallback generico pelo teclado
    overwrite_fast(None)


def press_tab(count: int = 1, pause: float = 0.12):
    for _ in range(count):
        send_keys("{TAB}")
        time.sleep(pause)


def wait_file_exists(path: str, timeout_s: float = 30.0) -> bool:
    def _ok():
        if not os.path.exists(path):
            return False
        try:
            with open(path, "rb"):
                return True
        except Exception:
            return False

    return wait_until(timeout_s, 0.25, _ok)


def find_similar_saved_file(save_dir: str, base_prefix: str) -> Optional[str]:
    try:
        entries = os.listdir(save_dir)
    except Exception:
        return None
    matches = []
    for name in entries:
        if name.lower().startswith(base_prefix.lower()):
            full = os.path.join(save_dir, name)
            try:
                mtime = os.path.getmtime(full)
            except Exception:
                mtime = 0
            matches.append((mtime, full))
    if not matches:
        return None
    matches.sort(reverse=True)
    return matches[0][1]


def ensure_saved_file(full_path: str, timeout_s: float = 45.0) -> str:
    if wait_file_exists(full_path, timeout_s=timeout_s):
        return full_path
    # tenta achar arquivo similar (talvez outra extensão adicionada pelo sistema)
    save_dir, base_name = os.path.split(full_path)
    base_prefix, _ = os.path.splitext(base_name)
    alt = find_similar_saved_file(save_dir, base_prefix)
    if alt and wait_file_exists(alt, timeout_s=5.0):
        log(f"[wait_file] encontrado arquivo similar: {alt}")
        return alt
    raise RuntimeError("Não confirmei o arquivo salvo no disco.")


def wait_for_saved_file(full_path: str, timeout_s: float, min_mtime: Optional[float] = None) -> Optional[str]:
    def _try_path(path: str) -> Optional[str]:
        if not os.path.exists(path):
            return None
        try:
            mtime = os.path.getmtime(path)
            if min_mtime is not None and mtime < min_mtime:
                return None
            with open(path, "rb"):
                return path
        except Exception:
            return None

    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        hit = _try_path(full_path)
        if hit:
            return hit
        save_dir, base_name = os.path.split(full_path)
        base_prefix, _ = os.path.splitext(base_name)
        alt = find_similar_saved_file(save_dir, base_prefix)
        if alt:
            hit = _try_path(alt)
            if hit:
                log(f"[wait_file] encontrado arquivo similar: {alt}")
                return hit
        time.sleep(0.25)
    return None


# =========================
# ATTACH / WINDOWS
# =========================

def attach_by_pid(pid: int) -> Optional[Application]:
    if not pid or pid <= 0:
        return None
    try:
        app = Application(backend="uia").connect(process=pid)
        return app
    except Exception:
        return None


def attach_by_title_contains(title_contains: str) -> Optional[Application]:
    target = normalize(title_contains)
    for p in psutil.process_iter(["pid", "name"]):
        pid = safe(lambda: p.info["pid"], None)
        if not pid:
            continue
        # Tenta achar alguma window do pid que contenha title_contains
        wins = iter_process_windows_uia(pid)
        for w in wins:
            txt = normalize(safe(lambda: w.window_text(), "") or "")
            if target and target in txt and txt:
                try:
                    return Application(backend="uia").connect(process=pid)
                except Exception:
                    pass
    return None


def wait_for_main_window(app: Application, timeout_s: float = 15.0) -> Optional[BaseWrapper]:
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        w = safe(lambda: app.top_window(), None)
        if w is not None:
            return w
        time.sleep(0.15)
    return None


def wait_for_window_containing_automation_id(pid: int, automation_id: str, timeout_s: float) -> Optional[BaseWrapper]:
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        for w in iter_process_windows_uia(pid):
            el = find_descendant_by_automation_id(w, automation_id)
            if el is not None:
                return w
        time.sleep(0.2)
    return None


def wait_for_report_window(pid: int, timeout_s: float = 60.0, extra_wait_s: float = 10.0) -> Optional[BaseWrapper]:
    # Primeiro, espera um pouco para a tela carregar
    time.sleep(extra_wait_s)

    w = wait_for_window_containing_automation_id(pid, AID_BTN_SALVAR_REL_EXCEL, timeout_s=timeout_s)
    if w:
        return w

    # Fallback por titulo + botao "Salvar"
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        for win in iter_process_windows_uia(pid):
            title = safe(lambda: win.window_text(), "") or ""
            nt = normalize(title)
            if "relatorio" in nt or "impressao" in nt:
                buttons = find_descendants(win, control_type="Button")
                for b in buttons:
                    bn = normalize(safe(lambda: b.window_text(), "") or "")
                    if "salvar" in bn:
                        return win
        time.sleep(0.2)
    return None


def wait_for_dre_window(pid: int, timeout_s: float = 40.0) -> Optional[BaseWrapper]:
    """
    Procura a janela/subjanela do DRE:
    - heuristica de titulo/campos/class_name (ex.: TfrmDRE, titulo "Demonstracao do Resultado do Exercicio").
    """
    targets = ["mês/ano", "mes/ano", "mes ano"]
    title_targets = ["demonstracao do resultado do exercicio", "demonstração do resultado do exercício"]
    class_targets = ["tfrmdre"]

    def has_mesano(win: BaseWrapper) -> bool:
        edits = find_descendants(win, control_type="Edit")
        for e in edits:
            name = normalize(safe(lambda: e.window_text(), "") or "")  # às vezes vazio
            name2 = normalize(safe(lambda: e.element_info.name, "") or "")
            if any(t in name or t in name2 for t in map(normalize, targets)):
                return True
        texts = find_descendants(win, control_type="Text")
        for t in texts:
            tn = normalize(safe(lambda: t.window_text(), "") or "")
            tn2 = normalize(safe(lambda: t.element_info.name, "") or "")
            if any(x in tn or x in tn2 for x in map(normalize, targets)):
                return True
        return False

    t0 = time.time()
    fallback = None
    seen_titles = set()
    while (time.time() - t0) < timeout_s:
        # tenta localizar contêiner pela classe/título típico
        dre = locate_dre_container(pid)
        if dre:
            return dre

        for win in iter_process_windows_uia(pid):
            title = normalize(safe(lambda: win.window_text(), "") or "")
            cls = normalize(safe(lambda: win.element_info.class_name, "") or "")
            if title:
                seen_titles.add(title)
            # também verifica sub-janelas de "Window"/"Pane"
            subwins = [win] + (safe(lambda: win.descendants(control_type="Window"), []) or []) + (safe(lambda: win.descendants(control_type="Pane"), []) or [])
            for sw in subwins:
                stitle = normalize(safe(lambda: sw.window_text(), "") or "")
                scls = normalize(safe(lambda: sw.element_info.class_name, "") or "")
                if stitle:
                    seen_titles.add(stitle)
                looks_like = (
                    ("dre" in stitle)
                    or any(t in stitle for t in title_targets)
                    or scls in class_targets
                    or cls in class_targets
                )
                if looks_like or has_mesano(sw):
                    return sw
                if fallback is None and ("demonstracao" in stitle or "dre" in stitle):
                    fallback = sw
        if fallback:
            return fallback
        time.sleep(0.2)

    if seen_titles:
        log(f"[wait_for_dre_window] titulos vistos: {sorted(seen_titles)}")
    return None


def is_dre_container(dre_win: Optional[BaseWrapper]) -> bool:
    if dre_win is None:
        return False
    if find_descendant_by_automation_id(dre_win, AID_BTN_IMPRIMIR):
        return True
    for aid in [AID_DRE_ESTAB_EDIT, AID_DRE_MES_INI_EDIT, AID_DRE_MES_FIM_EDIT, AID_DRE_GRAU_EDIT, AID_DRE_DATA_EMISSAO_EDIT]:
        if aid and find_descendant_by_automation_id(dre_win, aid):
            return True
    targets = ["mes/ano", "mes ano", "estabelecimento", "dre"]
    for ctrl in find_descendants(dre_win, control_type="Text") + find_descendants(dre_win, control_type="Edit"):
        nm = normalize(safe(lambda: ctrl.window_text(), "") or "") or normalize(safe(lambda: ctrl.element_info.name, "") or "")
        if any(t in nm for t in targets):
            return True
    return False


# =========================
# FINDERS (EMPRESA)
# =========================

def _record_coord(
    coords: dict,
    key: str,
    label: str,
    anchor_win: Optional[BaseWrapper],
    anchor_name: str,
):
    pos = capture_click_coordinates(label)
    if not pos:
        coords[key] = None
        return
    entry = {"x": int(pos[0]), "y": int(pos[1]), "anchor": anchor_name}
    if anchor_win:
        r = rect_of(anchor_win)
        if r:
            entry["dx"] = int(pos[0]) - r.left
            entry["dy"] = int(pos[1]) - r.top
    coords[key] = entry


def guided_capture_coords(main_win: BaseWrapper, coords_path: str):
    print("=== Modo guiado de coordenadas ===")
    print("Deixe o Contabil em primeiro plano e siga as instrucoes.")
    safe(lambda: main_win.set_focus(), None)
    time.sleep(0.2)

    def _anchor_name(win: Optional[BaseWrapper], fallback: str) -> str:
        if not win:
            return "main"
        if main_win:
            if safe(lambda: win.element_info.handle, None) == safe(lambda: main_win.element_info.handle, None):
                return "main"
        return fallback

    coords = {}
    _record_coord(coords, "empresa", "Empresa (campo onde digita)", main_win, "main")

    print("\nVou tentar abrir o DRE via ALT+R, N, R.")
    trigger_dre_shortcut()
    time.sleep(0.8)
    input("Se o DRE nao abriu, abra manualmente e pressione ENTER aqui...")

    pid = safe(lambda: main_win.element_info.process_id, None)
    dre_win = wait_for_dre_window(pid or 0, timeout_s=3.0) or main_win
    dre_anchor = _anchor_name(dre_win, "dre")
    for label, key in [
        ("Estabelecimento", "estab"),
        ("Mes/Ano inicial", "mes_ini"),
        ("Mes/Ano final", "mes_fim"),
        ("Grau", "grau"),
        ("Data de emissao", "data_emissao"),
        ("Botao Imprimir", "btn_imprimir"),
        ("Checkbox Listar Classificacao", "ckb_classificacao"),
        ("Checkbox Listar Titulos com Saldo Zero", "ckb_saldo_zero"),
        ("Checkbox Listar CNPJ", "ckb_cnpj"),
    ]:
        _record_coord(coords, key, label, dre_win, dre_anchor)

    print("\nAbra a tela de impressao (clique em Imprimir) e pressione ENTER aqui para capturar os botoes da pre-visualizacao.")
    input("Pressione ENTER quando a tela de impressao estiver visivel...")
    report_win = wait_for_report_window(pid or 0, timeout_s=5.0, extra_wait_s=0.1)
    if not report_win:
        report_win = safe(lambda: Desktop(backend="uia").get_active(), None) or main_win
    report_anchor = _anchor_name(report_win, "report")

    _record_coord(coords, "btn_salvar_excel", "Salvar Excel", report_win, report_anchor)
    _record_coord(coords, "fechar_impressao", "Fechar Impressao", report_win, report_anchor)

    save_coords_to_file(coords_path, coords)
    print(f"\n[coords] Coordenadas salvas em: {coords_path}")

def get_ordered_edits(root: BaseWrapper) -> List[BaseWrapper]:
    edits = find_descendants(root, control_type="Edit")
    scored = []
    for e in edits:
        r = rect_of(e)
        top = r.top if r else 10**9
        left = r.left if r else 10**9
        scored.append((top, left, e))
    scored.sort(key=lambda x: (x[0], x[1]))
    return [x[2] for x in scored]


def find_edit_by_label_proximity(root: BaseWrapper, edits: List[BaseWrapper], *labels: str) -> Optional[BaseWrapper]:
    targets = [normalize(x) for x in labels if x]
    # Considera qualquer control com texto/nome (não só Text), para cenários em que o label é Pane/Custom.
    labeled_controls = safe(lambda: root.descendants(), []) or []
    best = None
    best_score = 10**18

    for t in labeled_controls:
        tn = normalize(safe(lambda: t.window_text(), "") or "")
        tn2 = normalize(safe(lambda: t.element_info.name, "") or "")
        if not any(x and (x in tn or x in tn2) for x in targets):
            continue

        tr = rect_of(t)
        if not tr:
            continue

        for e in edits:
            er = rect_of(e)
            if not er:
                continue

            vdist = abs(er.top - tr.top)
            overlap = min(er.bottom, tr.bottom) - max(er.top, tr.top)
            if overlap < 0 and vdist > 60:
                continue

            # edit deve estar à direita do label
            if er.left < tr.right - 20:
                continue

            hspan = er.left - tr.right
            if hspan > 220:
                continue

            score = vdist + hspan
            if score < best_score:
                best_score = score
                best = e

    return best


def find_empresa_edit(root: BaseWrapper) -> Optional[BaseWrapper]:
    edits = find_descendants(root, control_type="Edit")

    # 1) por AutomationId direto
    for e in edits:
        aid = safe(lambda: e.element_info.automation_id, "") or ""
        if str(aid).strip().lower() == AID_EMPRESA_EDIT.lower():
            return e

    # 2) por label próximo
    by_label = find_edit_by_label_proximity(root, edits, "empresa")
    if by_label:
        return by_label

    # 3) pelo Name do próprio edit
    for e in edits:
        n = normalize(safe(lambda: e.element_info.name, "") or "")
        if "empresa" in n:
            return e

    # 4) heurística: primeiro edit no topo da janela
    ordered = get_ordered_edits(root)
    for e in ordered:
        r = rect_of(e)
        if r and r.top < 200:
            return e

    return None


def wait_for_empresa_edit(main_win: BaseWrapper, timeout_s: float = 4.0) -> Optional[BaseWrapper]:
    pid = safe(lambda: main_win.element_info.process_id, None)
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        tb = find_empresa_edit(main_win)
        if tb:
            return tb
        if pid:
            for w in iter_process_windows_uia(pid):
                tb = find_empresa_edit(w)
                if tb:
                    return tb
        time.sleep(0.2)
    return None


def find_control_by_automation_id_anywhere(pid: Optional[int], automation_id: str) -> Optional[BaseWrapper]:
    """Procura o AutomationId em todas as janelas do processo."""
    if not automation_id or not pid:
        return None
    for w in iter_process_windows_uia(pid):
        el = find_descendant_by_automation_id(w, automation_id)
        if el:
            return el
    return None


def find_control_by_automation_id_global(automation_id: str) -> Optional[BaseWrapper]:
    """Procura o AutomationId em qualquer janela da área de trabalho (qualquer processo)."""
    if not automation_id:
        return None
    desk = Desktop(backend="uia")
    wins = safe(lambda: desk.windows(), []) or []
    for w in wins:
        el = find_descendant_by_automation_id(w, automation_id)
        if el:
            return el
    return None


def find_edit_by_text(root: BaseWrapper, labels: List[str]) -> Optional[BaseWrapper]:
    """Procura um Edit cujo name/text contenha uma das labels normalizadas."""
    if not root:
        return None
    edits = find_descendants(root, control_type="Edit")
    targets = [normalize(x) for x in labels if x]
    for e in edits:
        nm = normalize(safe(lambda: e.window_text(), "") or "") or normalize(safe(lambda: e.element_info.name, "") or "")
        aid = normalize(str(safe(lambda: e.element_info.automation_id, "") or ""))
        for t in targets:
            if t and (t in nm or t == aid):
                return e
    return None


def debug_dump_windows(pid: Optional[int], target_aid: str):
    """Dump compacto de janelas/edits com AutomationId, para debug."""
    log(f"[debug_dump_windows] procurando AID={target_aid} em pid={pid}")
    wins = iter_process_windows_uia(pid or 0)
    for idx, w in enumerate(wins):
        title = safe(lambda: w.window_text(), "") or ""
        aid = safe(lambda: w.element_info.automation_id, "") or ""
        log(f"  [win {idx}] title='{title}' aid='{aid}'")
        edits = find_descendants(w, control_type="Edit")
        for e in edits:
            eaid = safe(lambda: e.element_info.automation_id, "") or ""
            enm = safe(lambda: e.element_info.name, "") or ""
            if target_aid and target_aid == str(eaid):
                log(f"    -> MATCH edit aid={eaid} name='{enm}'")
            else:
                log(f"    edit aid={eaid} name='{enm}'")


def locate_dre_container(pid: Optional[int]) -> Optional[BaseWrapper]:
    """
    Tenta localizar o contêiner do DRE procurando por classe ou título característico
    (ex.: classe 'TfrmDRE', título 'Demonstração do Resultado do Exercício').
    """
    title_targets = ["demonstracao do resultado do exercicio", "demonstração do resultado do exercício", "dre"]
    class_targets = ["tfrmdre"]

    def match(win: BaseWrapper) -> bool:
        t = normalize(safe(lambda: win.window_text(), "") or "")
        c = normalize(safe(lambda: win.element_info.class_name, "") or "")
        return any(x in t for x in title_targets) or c in class_targets

    for w in iter_process_windows_uia(pid or 0):
        if match(w):
            return w
        # procura em sub-janelas/panes
        subs = (safe(lambda: w.descendants(control_type="Window"), []) or []) + (safe(lambda: w.descendants(control_type="Pane"), []) or [])
        for sw in subs:
            if match(sw):
                return sw
    # tentativa global
    desk = Desktop(backend="uia")
    for w in safe(lambda: desk.windows(), []) or []:
        if match(w):
            return w
        subs = (safe(lambda: w.descendants(control_type="Window"), []) or []) + (safe(lambda: w.descendants(control_type="Pane"), []) or [])
        for sw in subs:
            if match(sw):
                return sw
    return None
# =========================
# AÇÕES (EQUIVALENTES AO C#)
# =========================

def set_empresa(main_win: BaseWrapper, empresa: int):
    safe(lambda: main_win.set_focus(), None)
    time.sleep(0.08)
    ensure_expected_active("set_empresa", main_win, require_active=False)

    if MANUAL_CLICK_MODE:
        prompt_user_click("Empresa (campo onde digita)")
        type_into_focused(empresa)
        send_keys("{ENTER}")
        time.sleep(0.1)
        send_keys("{ENTER}")
        time.sleep(0.3)
        return

    if COORD_EMPRESA:
        anchor = pick_anchor(COORD_EMPRESA, main_win=main_win)
        click_coord(COORD_EMPRESA, anchor, context="empresa")
        clear_focused_text()
        send_keys(str(empresa), with_spaces=True)
        time.sleep(FIELD_GAP_DELAY)
        send_keys("{ENTER}")
        time.sleep(0.1)
        send_keys("{ENTER}")
        time.sleep(0.3)
        return

    # igual ao C#: F2 e tentar achar o campo
    send_keys("{F2}")
    time.sleep(0.2)

    tb = wait_for_empresa_edit(main_win, timeout_s=4.0)
    if tb is None and not STRICT_EMPRESA_FIELD:
        ordered = get_ordered_edits(main_win)
        tb = ordered[0] if ordered else None

    if tb is None:
        raise RuntimeError("Campo de empresa nao encontrado; abortando para evitar cliques errados.")

    ensure_expected_active("set_empresa", main_win, require_active=False)
    replace_text(tb, str(empresa))
    time.sleep(FIELD_GAP_DELAY)
    send_keys("{ENTER}")
    time.sleep(0.1)
    # Extra ENTER para confirmar troca de empresa
    send_keys("{ENTER}")
    time.sleep(0.3)


def trigger_dre_shortcut():
    # equivalente ao C#: ALT+R, N, R (com delays)
    # em pywinauto: % = ALT
    # enviamos em etapas para manter parecido com o seu timing
    send_keys("%r")
    time.sleep(0.12)
    send_keys("n")
    time.sleep(0.12)
    send_keys("r")
    time.sleep(0.15)


def click_menu_item_by_id_or_name(roots: Iterable[BaseWrapper], automation_id: str, name_contains: str) -> bool:
    target_name = normalize(name_contains)
    for root in roots:
        items = find_descendants(root, control_type="MenuItem")
        for it in items:
            aid = str(safe(lambda: it.element_info.automation_id, "") or "").strip().lower()
            nm = normalize(safe(lambda: it.window_text(), "") or "") or normalize(safe(lambda: it.element_info.name, "") or "")
            if automation_id and aid == automation_id.lower():
                click_wrapper(it)
                return True
            if target_name and target_name in nm:
                click_wrapper(it)
                return True
    return False


def open_dre(main_win: BaseWrapper):
    safe(lambda: main_win.set_focus(), None)
    time.sleep(0.08)
    ensure_expected_active("open_dre", main_win)

    pid = safe(lambda: main_win.element_info.process_id, None)
    if not ASSUME_DRE_OPEN:
        # Se já houver janela do DRE aberta, reutiliza.
        existing = wait_for_dre_window(pid or 0, timeout_s=1.2)
        if existing:
            safe(lambda: existing.set_focus(), None)
            time.sleep(0.2)
            return

    trigger_dre_shortcut()
    time.sleep(0.5)

    roots = []
    if pid:
        roots = iter_process_windows_uia(pid)
    roots = roots or [main_win]

    if click_menu_item_by_id_or_name(roots, AID_DRE_MENU, "dre - demonstracao do resultado do exercicio"):
        time.sleep(0.5)
        return

    # fallback: tenta clicar o menu item pelo AID dentro da mainWin
    el = find_descendant_by_automation_id(main_win, AID_DRE_MENU)
    if el:
        click_wrapper(el)
        time.sleep(0.2)
        click_wrapper(el)
        time.sleep(0.7)


def configure_dre(
    dre_win: Optional[BaseWrapper],
    estab: int,
    periodo_ini: str,
    periodo_fim: str,
    do_checkboxes: bool,
    main_win: Optional[BaseWrapper] = None,
):
    base_anchor = pick_anchor(COORD_ESTAB, main_win=main_win, dre_win=dre_win)
    if not base_anchor:
        raise RuntimeError("Nenhuma janela ancora para coordenadas do DRE.")
    ensure_expected_active("configure_dre", base_anchor, require_active=False)
    if not is_dre_container(dre_win):
        log("[configure_dre] DRE container nao confirmado; usando modo por coordenadas.")
    safe(lambda: base_anchor.set_focus(), None)
    time.sleep(0.1)
    safe(lambda: base_anchor.click_input(), None)
    time.sleep(0.1)

    now = time.localtime()
    data_emissao = f"{now.tm_mday:02d}/{now.tm_mon:02d}/{now.tm_year:04d}"
    fields = [
        ("Estabelecimento", COORD_ESTAB, str(estab)),
        ("Mes/Ano inicial", COORD_MES_INI, periodo_ini),
        ("Mes/Ano final", COORD_MES_FIM, periodo_fim),
    ]
    if do_checkboxes:
        fields.extend([
            ("Grau", COORD_GRAU, str(GRAU)),
            ("Data emissao", COORD_DATA_EMISSAO, data_emissao),
        ])
    for label, pos, val in fields:
        if MANUAL_CLICK_MODE:
            prompt_user_click(label)
            type_into_focused(val)
        else:
            anchor = pick_anchor(pos, main_win=main_win, dre_win=dre_win)
            click_coord(pos, anchor or base_anchor, context=f"dre:{label}")
            replace_text(None, val, allow_foreign=True)
            time.sleep(FIELD_GAP_DELAY)

    if do_checkboxes:
        if MANUAL_CLICK_MODE:
            prompt_user_click("Radio Estabelecimento (selecione)")
            prompt_user_click("Checkbox Listar Classificacao (marcado)")
            prompt_user_click("Checkbox Listar Titulos com Saldo Zero (desmarcado)")
            prompt_user_click("Checkbox CNPJ (marcado)")
        else:
            radios = find_descendants(dre_win or base_anchor, control_type="RadioButton")
            for r in radios:
                rn = normalize(safe(lambda: r.window_text(), "") or "") or normalize(safe(lambda: r.element_info.name, "") or "")
                if "estabelecimento" in rn:
                    click_wrapper(r)
                    break

            if COORD_CKB_CLASSIFICACAO:
                anchor = pick_anchor(COORD_CKB_CLASSIFICACAO, main_win=main_win, dre_win=dre_win)
                click_coord(COORD_CKB_CLASSIFICACAO, anchor or base_anchor, context="ckb_classificacao")
            else:
                set_checkbox_by_label(dre_win or base_anchor, ["listar classificacao", "classificacao"], True)

            if COORD_CKB_TITULOS_SALDO_ZERO:
                anchor = pick_anchor(COORD_CKB_TITULOS_SALDO_ZERO, main_win=main_win, dre_win=dre_win)
                click_coord(COORD_CKB_TITULOS_SALDO_ZERO, anchor or base_anchor, context="ckb_saldo_zero")
            else:
                set_checkbox_by_label(dre_win or base_anchor, ["listar titulos com saldo zero"], False)

            if COORD_CKB_CNPJ:
                anchor = pick_anchor(COORD_CKB_CNPJ, main_win=main_win, dre_win=dre_win)
                click_coord(COORD_CKB_CNPJ, anchor or base_anchor, context="ckb_cnpj")
            else:
                set_checkbox_by_label(dre_win or base_anchor, ["cnpj"], True)


def click_by_automation_id_or_name(root: BaseWrapper, automation_id: str, alt_name_contains: Optional[List[str]] = None):
    el = find_descendant_by_automation_id(root, automation_id)
    if el:
        click_wrapper(el)
        return

    if alt_name_contains:
        buttons = find_descendants(root, control_type="Button")
        targets = [normalize(x) for x in alt_name_contains]
        for b in buttons:
            bn = normalize(safe(lambda: b.window_text(), "") or "") or normalize(safe(lambda: b.element_info.name, "") or "")
            if any(t in bn for t in targets):
                click_wrapper(b)
                return

    if automation_id == AID_BTN_IMPRIMIR and COORD_BTN_IMPRIMIR:
        click_coord(COORD_BTN_IMPRIMIR, root, context="btn_imprimir")
        return
    if automation_id == AID_BTN_SALVAR_REL_EXCEL and COORD_BTN_SALVAR_EXCEL:
        click_coord(COORD_BTN_SALVAR_EXCEL, root, context="btn_salvar_excel")
        return

    raise RuntimeError(f"Elemento não encontrado: {automation_id}")


def click_save_excel(
    report_win: Optional[BaseWrapper],
    main_win: Optional[BaseWrapper] = None,
    allow_unverified: bool = False,
) -> bool:
    if not report_win and COORD_BTN_SALVAR_EXCEL:
        report_win = find_container_by_offset(EXPECTED_PID, COORD_BTN_SALVAR_EXCEL, keywords=["salvar", "excel"])
    if report_win:
        safe(lambda: report_win.set_focus(), None)
        time.sleep(0.2)
        try:
            click_by_automation_id_or_name(
                report_win,
                AID_BTN_SALVAR_REL_EXCEL,
                alt_name_contains=["salvar relatorio excel", "salvar relatorio", "salvar relatorio", "salvar"],
            )
            return True
        except Exception:
            pass

    coord_abs = None
    if COORD_BTN_SALVAR_EXCEL:
        anchor = report_win if report_win else None
        coord_abs = resolve_coord(
            COORD_BTN_SALVAR_EXCEL,
            anchor_win=anchor,
            require_inside=False,
            context="btn_salvar_excel",
        )

    root = main_win or EXPECTED_MAIN_WIN or safe(lambda: Desktop(backend="uia").get_active(), None)
    if root:
        el = find_descendant_by_automation_id(root, AID_BTN_SALVAR_REL_EXCEL)
        if el:
            click_wrapper(el)
            return True
        if coord_abs:
            btn = find_button_near_coord(root, coord_abs, ["salvar", "excel", "relatorio"])
            if btn:
                click_wrapper(btn)
                return True

    if coord_abs and is_probable_save_target_at_point(coord_abs):
        click_abs(coord_abs, require_active=False)
        return True

    if coord_abs and allow_unverified:
        log("[save_click] usando coordenada sem confirmacao do botao.")
        click_abs(coord_abs, require_active=False)
        return True

    return False


def handle_save_dialog(full_path: str, timeout_s: float = 30.0) -> bool:
    desk = Desktop(backend="uia")
    save_dir, save_name = os.path.split(full_path)
    save_dir = save_dir or "."
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        wins = safe(lambda: desk.windows(), []) or []
        for w in wins:
            title = safe(lambda: w.window_text(), "") or ""
            if not title:
                continue
            nt = normalize(title)
            if "salvar" not in nt and "save" not in nt:
                continue

            safe(lambda: w.set_focus(), None)
            time.sleep(0.15)

            # força trocar a pasta via barra de endereço (Alt+D)
            send_keys("%d")
            time.sleep(0.15)
            replace_text(None, save_dir, allow_foreign=True)
            send_keys("{ENTER}")
            time.sleep(0.4)

            name_edit = find_descendant_by_automation_id(w, AID_SAVE_FILENAME)
            if name_edit:
                replace_text(name_edit, save_name, allow_foreign=True)
            else:
                replace_text(None, save_name, allow_foreign=True)

            save_btn = find_descendant_by_automation_id(w, AID_SAVE_BUTTON)
            if save_btn:
                click_wrapper(save_btn)
            else:
                # tenta Alt+S ou Enter como fallback
                send_keys("%s")
                time.sleep(0.2)
                send_keys("{ENTER}")

            time.sleep(0.3)
            return True

        time.sleep(0.15)

    return False


def handle_confirm_replace(timeout_s: float = 8.0):
    desk = Desktop(backend="uia")
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        wins = safe(lambda: desk.windows(), []) or []
        for w in wins:
            title = normalize(safe(lambda: w.window_text(), "") or "")
            if not title:
                continue

            if ("confirmar salvamento" in title) or ("confirm save as" in title) or ("confirm save" in title):
                safe(lambda: w.set_focus(), None)
                time.sleep(0.15)

                buttons = find_descendants(w, control_type="Button")
                for b in buttons:
                    bn = normalize(safe(lambda: b.window_text(), "") or "") or normalize(safe(lambda: b.element_info.name, "") or "")
                    if bn == "sim" or bn == "yes" or "substituir" in bn:
                        click_wrapper(b)
                        time.sleep(0.35)
                        return

                # fallback Enter
                send_keys("{ENTER}")
                time.sleep(0.35)
                return

        time.sleep(0.2)


def try_wait_saved_file(full_path: str, min_mtime: Optional[float], timeout_s: float = 6.0) -> Optional[str]:
    saved_path = wait_for_saved_file(full_path, timeout_s=timeout_s, min_mtime=min_mtime)
    if saved_path:
        return saved_path
    handle_confirm_replace(timeout_s=2.0)
    return wait_for_saved_file(full_path, timeout_s=timeout_s, min_mtime=min_mtime)


def close_report_by_coord(
    report_win: Optional[BaseWrapper],
    pid: Optional[int],
) -> bool:
    if not COORD_FECHAR_IMPRESSAO:
        return False
    anchor = report_win
    if not anchor and pid:
        anchor = find_container_by_offset(pid, COORD_FECHAR_IMPRESSAO, keywords=["fechar", "impressao", "relatorio"])
    if anchor:
        try:
            click_coord(COORD_FECHAR_IMPRESSAO, anchor, context="fechar_impressao")
            return True
        except Exception:
            pass
    pos = resolve_coord(COORD_FECHAR_IMPRESSAO, anchor_win=None, require_inside=False, context="fechar_impressao")
    if pos:
        try:
            click_abs(pos, require_active=False)
            return True
        except Exception:
            pass
    return False


def close_report_window(report_win: Optional[BaseWrapper], main_win: Optional[BaseWrapper] = None) -> bool:
    if MANUAL_CLICK_MODE:
        prompt_user_click("Fechar impressao/pre-visualizacao")
        return True
    if report_win is None:
        log("[close_report] report window not found; skip close.")
        return False

    if safe(lambda: report_win.exists(), True) is False:
        log("[close_report] report window no longer exists; skip close.")
        return False

    title = normalize(safe(lambda: report_win.window_text(), "") or "")
    if "relatorio" not in title and "impressao" not in title:
        log("[close_report] window title not report-like; skip close.")
        return False

    is_embedded = bool(main_win) and safe(lambda: report_win.element_info.handle, None) == safe(
        lambda: main_win.element_info.handle, None
    )
    win_rect = rect_of(report_win)
    save_btn = find_descendant_by_automation_id(report_win, AID_BTN_SALVAR_REL_EXCEL)
    save_rect = rect_of(save_btn) if save_btn else None

    def is_caption_close(btn_rect) -> bool:
        if not win_rect or not btn_rect:
            return False
        return btn_rect.top <= win_rect.top + 10 and btn_rect.right >= win_rect.right - 10

    def is_safe_close_button(btn: BaseWrapper) -> bool:
        btn_rect = rect_of(btn)
        if is_embedded:
            if is_caption_close(btn_rect):
                return False
            if not save_rect or not btn_rect:
                return False
            overlap = min(btn_rect.bottom, save_rect.bottom) - max(btn_rect.top, save_rect.top)
            if overlap < -5:
                return False
            if abs(btn_rect.top - save_rect.top) > 30:
                return False
        return True

    close_btn = find_descendant_by_automation_id(report_win, AID_BTN_FECHAR_PREV)
    if close_btn and is_safe_close_button(close_btn):
        click_wrapper(close_btn)
        time.sleep(CLICK_DELAY)
        return True

    buttons = find_descendants(report_win, control_type="Button")
    for b in buttons:
        bn = normalize(safe(lambda: b.window_text(), "") or "") or normalize(safe(lambda: b.element_info.name, "") or "")
        if "fechar" in bn or bn == "close":
            if not is_safe_close_button(b):
                continue
            click_wrapper(b)
            time.sleep(CLICK_DELAY)
            return True

    if not save_btn:
        log("[close_report] save-excel button not found; skip coord close.")
        return False

    if not COORD_FECHAR_IMPRESSAO:
        log("[close_report] close coord not configured; skip close.")
        return False

    if is_embedded:
        log("[close_report] embedded report; coord close disabled.")
        return False

    if not ALLOW_UNSAFE_COORD_CLOSE:
        log("[close_report] coord close disabled; skipping.")
        return False

    r = rect_of(report_win)
    if not r:
        log("[close_report] cannot read report window rect; skip close.")
        return False
    x, y = COORD_FECHAR_IMPRESSAO
    if not (r.left <= x <= r.right and r.top <= y <= r.bottom):
        log("[close_report] close coord outside report window; skip close.")
        return False

    safe(lambda: report_win.set_focus(), None)
    time.sleep(CLICK_DELAY)

    active = safe(lambda: Desktop(backend="uia").get_active(), None)
    if not active:
        log("[close_report] active window not found; skip close.")
        return False
    if safe(lambda: active.element_info.handle, None) != safe(lambda: report_win.element_info.handle, None):
        log("[close_report] report window not active; skip close.")
        return False

    pos = resolve_coord(COORD_FECHAR_IMPRESSAO, report_win, require_inside=True, context="fechar_impressao")
    if not pos:
        log("[close_report] close coord invalid; skip close.")
        return False
    click_abs(pos)
    time.sleep(CLICK_DELAY)
    return True


def dump_element_tree(root: BaseWrapper, max_depth: int = 5, depth: int = 0):
    if depth > max_depth:
        return

    indent = "  " * depth
    ctype = safe(lambda: root.friendly_class_name(), "") or ""
    name = safe(lambda: root.window_text(), "") or safe(lambda: root.element_info.name, "") or ""
    aid = safe(lambda: root.element_info.automation_id, "") or ""
    cls = safe(lambda: root.element_info.class_name, "") or ""
    print(f"{indent}{ctype} | Class={cls} | Name='{name}' | AutomationId='{aid}'")

    children = safe(lambda: root.children(), []) or []
    for ch in children:
        dump_element_tree(ch, max_depth=max_depth, depth=depth + 1)


def dump_focused():
    a = Application(backend="uia")
    el = safe(lambda: Desktop(backend="uia").get_active(), None)
    if not el:
        print("Nenhuma janela ativa.")
        return
    print("=== Janela ativa (aprox. focada) ===")
    dump_element_tree(el, max_depth=6)


# =========================
# MAIN
# =========================

def run(pid: int, save_root: str, title_contains: Optional[str] = None, app: Optional[Application] = None):
    os.makedirs(save_root, exist_ok=True)

    if app is None:
        title = title_contains or ATTACH_TITLE_CONTAINS
        app = attach_by_pid(pid) or attach_by_title_contains(title)
        if not app:
            raise RuntimeError(
                f"Não consegui anexar no PID {pid} nem pelo título '{title}'. Ajuste --pid ou --title-contains."
            )

    main_win = wait_for_main_window(app, timeout_s=15.0)
    if not main_win:
        raise RuntimeError("Não consegui localizar a janela principal do Contábil.")

    safe(lambda: main_win.set_focus(), None)
    time.sleep(0.2)

    proc_id = safe(lambda: main_win.element_info.process_id, pid)
    set_expected_context(proc_id, main_win)

    for emp in EMPRESAS:
        print(f"== Empresa {emp} ==")
        set_empresa(main_win, emp)

        estabs = ESTABS_721 if emp == 721 else [1]
        checkboxes_needed = True  # reconfigura checkboxes ao trocar de empresa
        for estab in estabs:
            print(f"-- Estab {estab} --")
            for periodo_ini, periodo_fim in RANGES:
                print(f"   Range {periodo_ini} -> {periodo_fim}")

                open_dre(main_win)

                dre_win = None
                dre_win = wait_for_dre_window(proc_id, timeout_s=2.0)
                if not ASSUME_DRE_OPEN:
                    dre_win = wait_for_dre_window(proc_id, timeout_s=40.0)
                    if not dre_win:
                        log("DRE não encontrado; tentando abrir novamente...")
                        open_dre(main_win)
                        dre_win = wait_for_dre_window(proc_id, timeout_s=40.0)
                    if not dre_win:
                        active = safe(lambda: Desktop(backend="uia").get_active(), None)
                        if active:
                            log("[dre_fallback] usando janela ativa como DRE.")
                            dre_win = active
                    if not dre_win:
                        raise RuntimeError("Não consegui localizar a janela do DRE (não achei o campo Mês/Ano).")
                if not dre_win:
                    dre_win = infer_anchor_from_coord(COORD_ESTAB) or main_win

                safe(lambda: dre_win.set_focus(), None)
                time.sleep(0.2)

                configure_dre(dre_win, estab, periodo_ini, periodo_fim, do_checkboxes=checkboxes_needed, main_win=main_win)
                checkboxes_needed = False

                # Imprimir (AutomationId ou coordenada)
                if MANUAL_CLICK_MODE:
                    prompt_user_click("Botao Imprimir")
                else:
                    dre_anchor = pick_anchor(COORD_BTN_IMPRIMIR, main_win=main_win, dre_win=dre_win)
                    click_by_automation_id_or_name(dre_anchor or dre_win, AID_BTN_IMPRIMIR, alt_name_contains=["imprimir"])

                # Aguarda alguns segundos apos imprimir e tenta salvar
                time.sleep(PRINT_WAIT_BEFORE_SAVE)
                report_win = None
                try:
                    report_win = wait_for_report_window(proc_id, timeout_s=5.0, extra_wait_s=0.0)
                except Exception:
                    report_win = None
                if not report_win and COORD_BTN_SALVAR_EXCEL:
                    report_win = find_container_by_offset(proc_id, COORD_BTN_SALVAR_EXCEL, keywords=["salvar", "excel"])
                    if report_win:
                        log("[report_fallback] container localizado via offset.")

                if MANUAL_CLICK_MODE:
                    prompt_user_click("Salvar Excel (na tela de impressao)")
                    save_clicked = True
                else:
                    if report_win:
                        safe(lambda: report_win.set_focus(), None)
                        time.sleep(0.3)
                    save_clicked = click_save_excel(report_win, main_win=main_win, allow_unverified=True)
                if not save_clicked:
                    raise RuntimeError("Nao consegui clicar no botao Salvar Excel.")
                save_attempt_t0 = time.time()

                perio_ini_safe = sanitize_filename(periodo_ini)
                perio_fim_safe = sanitize_filename(periodo_fim)
                file_name = f"{emp}.{estab}.{perio_ini_safe} a {perio_fim_safe}.xls"
                full_path = os.path.join(save_root, file_name)
                log(f"[save_target] {full_path}")

                if MANUAL_CLICK_MODE:
                    saved_path = wait_for_saved_file(full_path, timeout_s=45.0)
                else:
                    saved_path = try_wait_saved_file(full_path, min_mtime=save_attempt_t0 - 1.0, timeout_s=6.0)
                    if not saved_path:
                        if handle_save_dialog(full_path, timeout_s=20.0):
                            handle_confirm_replace(timeout_s=8.0)
                            saved_path = wait_for_saved_file(full_path, timeout_s=45.0, min_mtime=save_attempt_t0 - 1.0)
                        else:
                            log("[save_retry] Tentando clicar no botao de Excel novamente...")
                            click_save_excel(report_win, main_win=main_win, allow_unverified=True)
                            save_attempt_t0 = time.time()
                            saved_path = try_wait_saved_file(full_path, min_mtime=save_attempt_t0 - 1.0, timeout_s=6.0)
                            if not saved_path:
                                if handle_save_dialog(full_path, timeout_s=20.0):
                                    handle_confirm_replace(timeout_s=8.0)
                                    saved_path = wait_for_saved_file(full_path, timeout_s=45.0, min_mtime=save_attempt_t0 - 1.0)
                                else:
                                    saved_path = wait_for_saved_file(full_path, timeout_s=6.0, min_mtime=save_attempt_t0 - 1.0)

                if not saved_path:
                    raise RuntimeError("Falha ao salvar no dialogo do Windows.")
                log(f"[wait_file] confirmado arquivo: {saved_path}")

                closed = False
                if not MANUAL_CLICK_MODE:
                    closed = close_report_by_coord(report_win, proc_id)
                if not closed:
                    close_report_window(report_win, main_win)
                time.sleep(0.5)

    print("Concluído.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pid", type=int, default=ATTACH_PID, help="PID do Contábil para anexar (use 0 para ignorar)")
    ap.add_argument("--title-contains", type=str, default=ATTACH_TITLE_CONTAINS, help="Texto do título da janela do Contábil para anexar (fallback)")
    ap.add_argument("--save-root", type=str, default=SAVE_ROOT, help="Pasta de saída dos XLSX")
    ap.add_argument("--coords-file", type=str, default=COORDS_PATH, help="Arquivo JSON com coordenadas")
    ap.add_argument("--guided-coords", action="store_true", help="Modo guiado para capturar coordenadas e salvar no JSON")
    ap.add_argument("--manual-clicks", action="store_true", help="Modo manual: voce clica nos campos e o script so digita")
    ap.add_argument("--auto-clicks", action="store_true", help="Modo automatico: usar coordenadas quando disponiveis")
    ap.add_argument("--capture-coords", action="store_true", help="Modo captura: clique nos campos e confirme no console para registrar coordenadas")
    ap.add_argument("--dump", action="store_true", help="Dump simples da árvore da janela principal")
    ap.add_argument("--dump-focused", action="store_true", help="Dump da janela ativa (aprox. focada)")
    args = ap.parse_args()

    if args.manual_clicks:
        globals()["MANUAL_CLICK_MODE"] = True
    if args.auto_clicks:
        globals()["MANUAL_CLICK_MODE"] = False

    if args.dump_focused:
        dump_focused()
        return

    app = attach_by_pid(args.pid) or attach_by_title_contains(args.title_contains)
    if not app:
        raise SystemExit(f"Não consegui anexar no PID {args.pid} nem pelo título '{args.title_contains}'. Ajuste --pid ou --title-contains.")

    main_win = wait_for_main_window(app, timeout_s=15.0)
    if not main_win:
        raise SystemExit("Não consegui localizar a janela principal.")

    load_coords_from_file(args.coords_file)

    if args.dump:
        print("=== Janela Principal ===")
        dump_element_tree(main_win, max_depth=7)
        return

    if args.guided_coords:
        guided_capture_coords(main_win, args.coords_file)
        return

    if args.capture_coords:
        print("Modo captura de coordenadas: aperte ENTER e depois clique no campo para cada label.")
        for lbl in [
            "Estabelecimento",
            "Mes/Ano inicial",
            "Mes/Ano final",
            "Grau",
            "Data de emissao",
            "Botao Imprimir",
            "Salvar Excel",
            "Checkbox Listar Classificacao",
            "Checkbox Listar Titulos com Saldo Zero",
            "Checkbox Listar CNPJ",
            "Fechar Impressao",
        ]:
            capture_click_coordinates(lbl)
        return

    run(args.pid, args.save_root, title_contains=args.title_contains, app=app)


if __name__ == "__main__":
    main()
