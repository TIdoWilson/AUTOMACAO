import sys
import re
import csv
import json
import os
import time
import ctypes
import logging
from ctypes import wintypes
from datetime import datetime
from typing import Any, Dict, Optional, Tuple, List

import pyautogui
import uiautomation as auto

# --------------------- configuracoes ---------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
LOG_DIR = os.path.join(BASE_DIR, "logs")
LOG_FILE = os.path.join(LOG_DIR, "informe_teste.log")
ERROR_LOG_FILE = os.path.join(LOG_DIR, "informe_teste_erros.txt")

# Macros gravados (locators/points)
MACRO_PATH = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\5 - Informe de rendimentos\Macros\macro_20260128_084041.json"
MINI_MACRO_PATH = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\5 - Informe de rendimentos\Macros\macro_20260128_085432.json"
SAVE_PDF_MACRO_PATH = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\5 - Informe de rendimentos\Macros\macro_20260128_085844.json"
CLOSE_PRINT_WINDOW_MACRO_PATH = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\5 - Informe de rendimentos\Macros\macro_20260128_090735.json"

# CSV de empresas/estabelecimentos (padrao)
CSV_EMPRESAS = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\5 - Informe de rendimentos\empresas.csv"

# Janela principal
WIN_MAIN_NAME = "Folha de Pagamento "
WIN_MAIN_CLASS = "TfrmPrincipal"

# Valores dinamicos
empresa_sistema = "empresa_sistema"
codigo_empresa_excel = "177"
DATA_ATUAL = datetime.now().strftime("%d%m%Y")
ANO_BASE = str(datetime.now().year - 1)
responsavel = "1"
classificacao = "1"
estabelecimento_empresa_atual = ""  # PREENCHER
PASTA_DESTINO_PDF = r"\\192.0.0.251\Arquivos\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\5 - Informe de rendimentos\Relatórios\2026"

# Steps do macro base (ids)
STEP_CODIGO = 2
STEP_MENU_ABRIR = 47
STEP_MENU_ITEM_1 = 49
STEP_MENU_ITEM_2 = 51
STEP_DATA = 26
STEP_ANO = 53
STEP_CAMPO1 = 63
STEP_CAMPO2 = 71
STEP_DISPENSA = 77
STEP_GERAL = 78

# Steps do mini macro (ids)
STEP_ESTABELECIMENTO = 2
STEP_CONFIRMAR = 6
STEP_IMPRIMIR = 8

# Steps do salvar PDF (ids)
STEP_ABRIR_SALVAR_PDF = 2
STEP_ABRIR_SELECAO_DIRETORIO = 4
STEP_SELECIONAR_PASTA = 19
STEP_CAMPO_NOME_ARQUIVO = 21
STEP_CHECK_ABRIR_DOC = 29
STEP_CHECK_PDFA = 31
STEP_OK_SALVAR = 33
STEP_OK_AVISO = 35

# Steps fechar janela extra (ids)
STEP_FECHAR_JANELA_EXTRA = 2

# Opcional: usar ponto para estabelecimento/imprimir se nao achar via UIA
PONTO_ESTABELECIMENTO = None  # ex.: (0.50, 0.60)
PONTO_IMPRIMIR = None  # ex.: (0.90, 0.85)

# --------------------- Win32 helpers ---------------------
user32 = ctypes.windll.user32

user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
user32.FindWindowW.restype = wintypes.HWND

user32.EnumWindows.argtypes = [ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM), wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL

user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int

user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int

user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.GetWindowRect.restype = wintypes.BOOL

user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
user32.ShowWindow.restype = wintypes.BOOL

user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.SetForegroundWindow.restype = wintypes.BOOL

user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND

SW_RESTORE = 9
SW_MAXIMIZE = 3

# --------------------- logger ---------------------
def setup_logger() -> logging.Logger:
    os.makedirs(LOG_DIR, exist_ok=True)
    logger = logging.getLogger("informe_teste")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
    sh = logging.StreamHandler(); sh.setFormatter(fmt); logger.addHandler(sh)
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8"); fh.setFormatter(fmt); logger.addHandler(fh)
    return logger

def log_erro(codigo: str, estab: str, motivo: str):
    os.makedirs(LOG_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts};{codigo};{estab};{motivo}\n"
    with open(ERROR_LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line)

# --------------------- util ---------------------
def get_text(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(wintypes.HWND(hwnd), buf, 512)
    return buf.value

def get_class(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(wintypes.HWND(hwnd), buf, 256)
    return buf.value

def get_rect(hwnd: int) -> Optional[Tuple[int, int, int, int]]:
    r = wintypes.RECT()
    if not user32.GetWindowRect(wintypes.HWND(hwnd), ctypes.byref(r)):
        return None
    return (int(r.left), int(r.top), int(r.right), int(r.bottom))

def enum_windows() -> List[int]:
    out: List[int] = []
    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def cb(hwnd, lparam):
        out.append(int(hwnd))
        return True
    user32.EnumWindows(cb, 0)
    return out

def find_window_fuzzy(name: Optional[str], cls: Optional[str]) -> int:
    if name is not None or cls is not None:
        h = user32.FindWindowW(cls, name)
        if h:
            return int(h)
    want_name = (name or "").strip().lower()
    want_cls = (cls or "").strip()
    for h in enum_windows():
        got_name = get_text(h)
        got_cls = get_class(h)
        if want_cls and got_cls != want_cls:
            continue
        if want_name and want_name not in (got_name or "").lower():
            continue
        return h
    return 0

def focus_window(name: str, cls: str, logger: logging.Logger, mode: str = "keep") -> int:
    hwnd = find_window_fuzzy(name, cls)
    if not hwnd:
        raise RuntimeError(f"Janela nao encontrada: name={name} class={cls}")
    if mode == "maximize":
        user32.ShowWindow(wintypes.HWND(hwnd), SW_MAXIMIZE)
    elif mode == "restore":
        user32.ShowWindow(wintypes.HWND(hwnd), SW_RESTORE)
    user32.SetForegroundWindow(wintypes.HWND(hwnd))
    logger.info("Foco na janela: %s (%s)", name, cls)
    time.sleep(0.2)
    return hwnd

def click_window_point(hwnd: int, rx: float, ry: float, button: str = "left") -> Tuple[int, int]:
    rect = get_rect(hwnd)
    if not rect:
        raise RuntimeError("Sem rect da janela")
    l, t, r, b = rect
    x = int(l + (r - l) * rx)
    y = int(t + (b - t) * ry)
    pyautogui.click(x, y, button=button)
    return x, y

def press_backspace(n: int = 5):
    pyautogui.press("backspace", presses=n)

def press_delete(n: int = 10):
    pyautogui.press("delete", presses=n)

def type_text(text: str):
    if text:
        pyautogui.write(text, interval=0.0)

def type_text_logged(text: str, logger: logging.Logger, label: str):
    logger.info("Digite %s: %s", label, text)
    type_text(text)

def press_enter(n: int = 1):
    pyautogui.press("enter", presses=n)

def set_clipboard_text(text: str):
    try:
        import pyperclip
        pyperclip.copy(text)
    except Exception:
        pass

def parse_estabelecimentos(value: str) -> List[str]:
    if not value:
        return []
    parts = re.split(r"[;,/ ]+", value.strip())
    out = []
    for part in parts:
        if not part:
            continue
        out.append(part)
    return out

def load_empresas_csv(path: str) -> List[Dict[str, Any]]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"CSV nao encontrado: {path}")
    rows: List[Dict[str, Any]] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        header = next(reader, None)
        if not header:
            return rows
        for row in reader:
            if not row:
                continue
            codigo = row[0].strip() if len(row) >= 1 else ""
            nome = row[1].strip() if len(row) >= 2 else ""
            estab_raw = row[2].strip() if len(row) >= 3 else ""
            if not codigo:
                continue
            estabs = parse_estabelecimentos(estab_raw)
            if not estabs:
                estabs = ["1"]
            rows.append({"codigo": codigo, "nome": nome or codigo, "estabelecimentos": estabs})
    return rows

def sanitize_filename(name: str) -> str:
    invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid:
        name = name.replace(ch, "-")
    return " ".join(name.split()).strip()

def build_nome_arquivo(nome: str, estab: str) -> str:
    if str(estab).strip() == "1":
        base = f"INFORME - {nome}"
    else:
        base = f"INFORME - {nome} - {estab}"
    return sanitize_filename(base)

# --------------------- UIA helpers ---------------------
def normalize_control_type(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    s = str(s)
    return s[:-7] if s.endswith("Control") else s

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

def resolve_by_locator_tree(uia_win, locator: Dict[str, Any], max_depth: int = 40):
    for ctrl, _ in iter_descendants(uia_win, max_depth):
        try:
            if is_match(ctrl, locator) and ancestors_ok(ctrl, locator.get("ancestors", [])):
                return ctrl
        except Exception:
            continue
    return None

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

def find_uia_control(win, control_type: Optional[str], name: Optional[str], class_name: Optional[str], search_depth: int = 40):
    locator = {"control_type": control_type, "name": name, "class_name": class_name, "ancestors": []}
    return resolve_by_locator_tree(win, locator, max_depth=search_depth)

def click_uia_named(win, control_type: Optional[str], name: Optional[str], class_name: Optional[str], logger: logging.Logger):
    ctrl = find_uia_control(win, control_type, name, class_name)
    if ctrl:
        try:
            ctrl.Click()
            return
        except Exception:
            pass
    raise RuntimeError(f"Controle nao encontrado: type={control_type} name={name} class={class_name}")

def wait_dialog(name: str, class_name: str, timeout_s: float = 8.0) -> bool:
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        dlg = auto.WindowControl(Name=name, ClassName=class_name)
        if dlg.Exists(0.2, 0.05):
            return True
        time.sleep(0.1)
    return False

def wait_dialog_close(name: str, class_name: str, timeout_s: float = 8.0) -> bool:
    t0 = time.time()
    while (time.time() - t0) < timeout_s:
        dlg = auto.WindowControl(Name=name, ClassName=class_name)
        if not dlg.Exists(0.2, 0.05):
            return True
        time.sleep(0.1)
    return False

def handle_aviso_sem_dados(timeout_s: float = 4.0) -> bool:
    if not wait_dialog("Aviso do Sistema", "#32770", timeout_s=timeout_s):
        return False
    dlg = auto.WindowControl(Name="Aviso do Sistema", ClassName="#32770")
    try:
        btn = dlg.ButtonControl(Name="OK")
        if btn.Exists(0.2, 0.05):
            btn.Click()
        else:
            pyautogui.press("enter")
    except Exception:
        pyautogui.press("enter")
    wait_dialog_close("Aviso do Sistema", "#32770", timeout_s=timeout_s)
    return True

def handle_sobrescrever_pdf(timeout_s: float = 6.0) -> bool:
    dlg = auto.WindowControl(Name="Aviso do Sistema", ClassName="#32770")
    if not dlg.Exists(timeout_s, 0.1):
        return False
    try:
        btn = dlg.ButtonControl(Name="Sim")
        if btn.Exists(0.2, 0.05):
            btn.Click()
            wait_dialog_close("Aviso do Sistema", "#32770", timeout_s=timeout_s)
            return True
    except Exception:
        pass
    return False
def click_uia_step(step: Dict[str, Any], logger: logging.Logger):
    target = step["target"]
    win_spec = target["window"]
    locator = target["locator"]
    logger.info("Clique UIA: janela=%r classe=%r controle=%r nome=%r",
                win_spec.get("name"), win_spec.get("class_name"),
                locator.get("control_type"), locator.get("name"))

    # Menu popup (#32768) nao tem nome: procurar por classe sem forcar foco
    if (win_spec.get("class_name") == "#32768") and not (win_spec.get("name") or ""):
        hwnd = 0
        for h in enum_windows():
            if get_class(h) == "#32768":
                hwnd = h
                break
        if not hwnd:
            hwnd = int(user32.GetForegroundWindow() or 0)
    else:
        hwnd = focus_window(win_spec.get("name", ""), win_spec.get("class_name", ""), logger, mode="keep")
    uia_win = get_uia_window_from_hwnd(hwnd, win_spec.get("class_name"), win_spec.get("name"))

    resolved = None
    if uia_win:
        resolved = resolve_by_locator_tree(uia_win, locator, max_depth=int(locator.get("search_depth") or 40))

    if resolved:
        rect = getattr(resolved, "BoundingRectangle", None)
        if rect:
            l, t, r, b = int(rect.left), int(rect.top), int(rect.right), int(rect.bottom)
            rx = step["point"]["rx"]
            ry = step["point"]["ry"]
            x = int(l + (r - l) * rx)
            y = int(t + (b - t) * ry)
            pyautogui.click(x, y, button=step.get("button", "left"))
            return

    # Fallback para ponto relativo à janela
    wp = step.get("window_point") or step.get("point")
    click_window_point(hwnd, wp["rx"], wp["ry"], step.get("button", "left"))

def click_win_step(step: Dict[str, Any], logger: logging.Logger):
    win_spec = step["window"]
    hwnd = focus_window(win_spec.get("name", ""), win_spec.get("class_name", ""), logger, mode="keep")
    pt = step["point"]
    logger.info("Clique WIN: janela=%r classe=%r ponto=%s", win_spec.get("name"), win_spec.get("class_name"), pt)
    click_window_point(hwnd, pt["rx"], pt["ry"], step.get("button", "left"))

def click_step_window_point(step: Dict[str, Any], logger: logging.Logger):
    target = step.get("target") or {}
    win_spec = target.get("window") or {}
    hwnd = focus_window(win_spec.get("name", ""), win_spec.get("class_name", ""), logger, mode="keep")
    pt = step.get("window_point") or step.get("point")
    logger.info("Clique por coordenada: janela=%r classe=%r ponto=%s", win_spec.get("name"), win_spec.get("class_name"), pt)
    click_window_point(hwnd, pt["rx"], pt["ry"], step.get("button", "left"))

def load_macro_steps(path: str) -> Dict[int, Dict[str, Any]]:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    steps = data.get("steps", [])
    out = {}
    for s in steps:
        sid = int(s.get("id", 0))
        if sid:
            out[sid] = s
    return out

# --------------------- fluxo ---------------------
def abrir_menu_informe(hwnd_main: int, logger: logging.Logger):
    del hwnd_main
    del logger
    raise RuntimeError("Use steps do macro para abrir o menu.")

def preencher_estabelecimento(win_inf, hwnd_inf: int, logger: logging.Logger):
    if PONTO_ESTABELECIMENTO:
        click_window_point(hwnd_inf, PONTO_ESTABELECIMENTO[0], PONTO_ESTABELECIMENTO[1])
        press_backspace(5)
        type_text(ESTABELECIMENTO)
        logger.info("Estabelecimento preenchido (ponto)")
        return

    ctrl = find_uia_control(win_inf, "Edit", "Estabelecimento", None)
    if ctrl:
        ctrl.Click()
        press_backspace(5)
        type_text(ESTABELECIMENTO)
        logger.info("Estabelecimento preenchido (UIA)")
        return

    raise RuntimeError("Nao foi possivel localizar o campo Estabelecimento. Ajuste PONTO_ESTABELECIMENTO ou o nome.")

def clicar_imprimir(win_inf, hwnd_inf: int, logger: logging.Logger):
    if PONTO_IMPRIMIR:
        click_window_point(hwnd_inf, PONTO_IMPRIMIR[0], PONTO_IMPRIMIR[1])
        logger.info("Imprimir clicado (ponto)")
        return

    ctrl = find_uia_control(win_inf, "Button", "Imprimir", None)
    if ctrl:
        ctrl.Click()
        logger.info("Imprimir clicado (UIA)")
        return

    raise RuntimeError("Nao foi possivel localizar o botao Imprimir. Ajuste PONTO_IMPRIMIR ou o nome.")

def main():
    logger = setup_logger()
    pyautogui.FAILSAFE = True

    steps = load_macro_steps(MACRO_PATH)
    mini_steps = load_macro_steps(MINI_MACRO_PATH)
    save_steps = load_macro_steps(SAVE_PDF_MACRO_PATH)
    close_steps = load_macro_steps(CLOSE_PRINT_WINDOW_MACRO_PATH)

    csv_path = sys.argv[1] if len(sys.argv) > 1 else CSV_EMPRESAS
    empresas = load_empresas_csv(csv_path)
    if not empresas:
        logger.info("Nenhuma empresa encontrada no CSV.")
        return

    # Maximizar apenas a janela principal uma vez no inicio
    focus_window(WIN_MAIN_NAME, WIN_MAIN_CLASS, logger, mode="maximize")

    for empresa in empresas:
        codigo = empresa["codigo"]
        nome = empresa.get("nome") or codigo
        estabs = empresa.get("estabelecimentos") or ["1"]

        logger.info("=== Empresa %s | %s ===", codigo, nome)

        # Campo codigo (barra superior)
        click_step_window_point(steps[STEP_CODIGO], logger)
        press_backspace(5)
        type_text_logged(codigo, logger, "codigo_empresa")
        press_enter(2)
        time.sleep(0.5)

        # Campo data (hoje)
        click_step_window_point(steps[STEP_DATA], logger)
        press_backspace(10)
        press_delete(10)
        type_text_logged(DATA_ATUAL, logger, "data_sistema")
        press_enter(2)
        time.sleep(0.5)

        # Abrir menu do informe
        click_win_step(steps[STEP_MENU_ABRIR], logger)
        time.sleep(0.3)
        click_uia_step(steps[STEP_MENU_ITEM_1], logger)
        time.sleep(0.3)
        click_uia_step(steps[STEP_MENU_ITEM_2], logger)

        # Foco na janela principal (MDI)
        hwnd_main = focus_window(WIN_MAIN_NAME, WIN_MAIN_CLASS, logger)
        win_main = get_uia_window_from_hwnd(hwnd_main, WIN_MAIN_CLASS, WIN_MAIN_NAME)
        if not win_main:
            raise RuntimeError("Janela principal nao encontrada via UIA")

        # Data do sistema (reforco antes dos campos 1/1)
        click_step_window_point(steps[STEP_DATA], logger)
        time.sleep(0.3)
        press_backspace(10)
        press_delete(10)
        type_text(DATA_ATUAL)
        press_enter(2)
        time.sleep(0.5)

        # Campos com 1 (usar coordenada)
        click_step_window_point(steps[STEP_CAMPO1], logger)
        time.sleep(0.2)
        type_text_logged(responsavel, logger, "responsavel")
        press_enter(1)
        time.sleep(0.3)

        click_step_window_point(steps[STEP_CAMPO2], logger)
        time.sleep(0.2)
        type_text_logged(classificacao, logger, "classificacao")
        press_enter(1)
        time.sleep(0.3)

        # Marcar dispensa de assinatura
        step_dispensa = steps.get(STEP_DISPENSA)
        if step_dispensa and "target" in step_dispensa:
            click_uia_step(step_dispensa, logger)
        else:
            click_uia_named(win_main, "CheckBox", "Dispensa de assinatura", None, logger)

        # Marcar geral
        step_geral = steps.get(STEP_GERAL)
        if step_geral and "target" in step_geral:
            click_uia_step(step_geral, logger)
        else:
            click_uia_named(win_main, "RadioButton", "Geral", None, logger)

        for estab in estabs:
            logger.info("--- Estabelecimento %s ---", estab)
            # Preencher estabelecimento (mini macro)
            click_step_window_point(mini_steps[STEP_ESTABELECIMENTO], logger)
            press_backspace(5)
            type_text_logged(estab, logger, "estabelecimento")

            # Confirmar (mini macro)
            click_step_window_point(mini_steps[STEP_CONFIRMAR], logger)

            # Imprimir para abrir a proxima parte (mini macro)
            click_uia_step(mini_steps[STEP_IMPRIMIR], logger)
            time.sleep(0.7)

            # Se aparecer aviso de "nao ha dados", confirma e vai para o proximo
            if handle_aviso_sem_dados(timeout_s=6.0):
                log_erro(codigo, estab, "Aviso do Sistema: nao ha dados para listar")
                continue

            # Salvar PDF
            try:
                click_step_window_point(save_steps[STEP_ABRIR_SALVAR_PDF], logger)
                time.sleep(0.5)
                click_step_window_point(save_steps[STEP_ABRIR_SELECAO_DIRETORIO], logger)

                # Selecionar pasta direto pelo caminho
                focus_window("Seleção de Diretório", "#32770", logger, mode="keep")
                time.sleep(0.3)
                pyautogui.hotkey("alt", "d")
                time.sleep(0.2)
                pyautogui.hotkey("ctrl", "a")
                set_clipboard_text(PASTA_DESTINO_PDF)
                pyautogui.hotkey("ctrl", "v")
                press_enter(1)
                time.sleep(0.6)
                click_uia_step(save_steps[STEP_SELECIONAR_PASTA], logger)

                time.sleep(0.5)
                click_step_window_point(save_steps[STEP_CAMPO_NOME_ARQUIVO], logger)
                pyautogui.hotkey("ctrl", "a")
                type_text_logged(build_nome_arquivo(nome, estab) + ".pdf", logger, "nome_pdf")

                click_uia_step(save_steps[STEP_CHECK_ABRIR_DOC], logger)
                click_uia_step(save_steps[STEP_CHECK_PDFA], logger)
                click_uia_step(save_steps[STEP_OK_SALVAR], logger)
                handle_sobrescrever_pdf(timeout_s=6.0)
                if wait_dialog("Aviso do Sistema", "#32770", timeout_s=12.0):
                    click_uia_step(save_steps[STEP_OK_AVISO], logger)
                    wait_dialog_close("Aviso do Sistema", "#32770", timeout_s=12.0)

                # Fechar janela extra de impressao
                click_step_window_point(close_steps[STEP_FECHAR_JANELA_EXTRA], logger)
                time.sleep(0.5)
            except Exception as exc:
                log_erro(codigo, estab, f"Erro ao salvar PDF: {exc}")
                continue

        # Voltar para selecao de empresa
        click_step_window_point(steps[STEP_CODIGO], logger)
        time.sleep(0.5)

    logger.info("Fluxo concluido")

if __name__ == "__main__":
    main()
