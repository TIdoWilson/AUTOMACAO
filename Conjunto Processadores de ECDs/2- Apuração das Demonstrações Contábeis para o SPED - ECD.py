# -*- coding: utf-8 -*-
"""
Apuração das Demonstrações Contábeis para o SPED - ECD
"""

from __future__ import annotations

import datetime as _dt
import os
import sys
import time

try:
    import pyautogui as pag
except Exception as exc:
    raise SystemExit("pyautogui é obrigatório para a automação de UI") from exc

try:
    from pywinauto import Desktop
    from pywinauto import findwindows
    from pywinauto.base_wrapper import BaseWrapper
except Exception as exc:
    raise SystemExit("pywinauto é obrigatório para ler/marcar checkboxes") from exc

# =====================
# Configuracoes
# =====================
NOME_SCRIPT = "2 - Apuração das Demonstrações Contábeis para o SPED - ECD"
BASE_DIR = r"W:\\SPEDs\\ECD\\2025"
ANO_PASSADO = _dt.date.today().year - 1
LOG_PATH = os.path.join(os.path.dirname(__file__), "log_apuracao_demonstracoes_ecd.txt")

# Coordenadas
BTN_OK = (1172, 419)
BTN_POS_PROCESSO_1 = (1042, 598)
BTN_POS_PROCESSO_2 = (1218, 382)

# Tempos
WAIT_START = 2
WAIT_APOS_OK = 120
WAIT_CURTO = 0.5
WAIT_JANELA = 8
DELAY_APOS_BIND = 0.1
DELAY_ENTRE_CLIQUES = 0.01

AVISO_TITULO = "Aviso do Sistema"
AVISO_CLASS = "#32770"
AVISO_OK_CLASS = "Button1"

# UIA/Win32
CHECK_MENSAL_CLASS = "TCRDCheckBox7"
CHECK_ANUAL_CLASS = "TCRDCheckBox5"
CHECK_GERAR_SALDO_ANTERIOR_CLASS = "TCRDCheckBox2"
CHECK_EXCLUIR_DEMONSTRACOES_CLASS = "TCRDCheckBox3"
CHECK_GERAR_CONTAS_SALDO_ZERO_CLASS = "TCRDCheckBox"
CHECK_ABERTURA_EMPRESA_DIFERENTE_CLASS = "TCRDCheckBox"
CHECK_TRIMESTRAL_CLASS = "TCRDCheckBox"
CHECK_SEMESTRAL_CLASS = "TCRDCheckBox"

# pyautogui
pag.PAUSE = 0.05
pag.FAILSAFE = False


# =====================
# Utilitarios
# =====================

def _log(msg: str) -> None:
    timestamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    linha = f"[{timestamp}] {msg}"
    print(linha)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(linha + "\n")


def _obter_empresa(argv: list[str]) -> str:
    if len(argv) > 1:
        partes: list[str] = []
        for a in argv[1:]:
            if a.strip().startswith("--"):
                break
            partes.append(a)
        if partes:
            return " ".join(partes).strip()
    return input("Empresa (nome completo): ").strip()

def _modo_sem_ok(argv: list[str]) -> bool:
    return "--sem-ok" in [a.strip().lower() for a in argv]


def _caminho_destino(empresa: str) -> str:
    return f"{BASE_DIR}\\{empresa}"


def _sleep(segundos: float, motivo: str | None = None) -> None:
    if motivo:
        _log(f"Aguardando {segundos}s: {motivo}")
    time.sleep(segundos)


def _alt_sequencia(teclas: str) -> None:
    pag.keyDown("alt")
    for t in teclas:
        pag.press(t)
    pag.keyUp("alt")


def _click(pos: tuple[int, int], motivo: str) -> None:
    _log(f"Clique: {motivo} {pos}")
    pag.click(pos[0], pos[1])

def _normalizar(txt: str) -> str:
    return (txt or "").strip().lower()

def _achar_checkbox_por_texto(texto: str) -> BaseWrapper | None:
    try:
        elementos = findwindows.find_elements(
            backend="win32",
            class_name_re=r"TCRDCheckBox.*",
            top_level_only=False,
        )
        for el in elementos:
            if _normalizar(el.name) == _normalizar(texto) or _normalizar(texto) in _normalizar(el.name):
                _log(f"Checkbox por texto localizado: '{el.name}' ({el.class_name})")
                return Desktop(backend="win32").window(handle=el.handle).wrapper_object()
    except Exception:
        return None
    return None


def _achar_checkbox_por_class(class_name: str) -> BaseWrapper | None:
    try:
        elementos = findwindows.find_elements(
            backend="win32",
            class_name=class_name,
            top_level_only=False,
        )
        if elementos:
            el = elementos[0]
            _log(f"Checkbox por class localizado: {class_name} texto='{el.name}'")
            return Desktop(backend="win32").window(handle=el.handle).wrapper_object()
    except Exception:
        return None
    return None


def _log_checkboxes_disponiveis() -> None:
    try:
        elementos = findwindows.find_elements(
            backend="win32",
            class_name_re=r"TCRDCheckBox.*",
            top_level_only=False,
        )
        textos = [f"{el.class_name}='{el.name}'" for el in elementos if el.name]
        if textos:
            _log("Checkboxes detectados: " + " | ".join(textos))
    except Exception:
        pass


def _obter_checkbox(class_name: str, texto: str) -> BaseWrapper | None:
    ctrl = _achar_checkbox_por_texto(texto)
    if ctrl is not None:
        return ctrl
    ctrl = _achar_checkbox_por_class(class_name)
    if ctrl is not None:
        return ctrl
    return None

def _check_state(ctrl) -> int | None:
    try:
        return int(ctrl.get_check_state())
    except Exception:
        return None

def _garantir_checkbox(class_name: str, texto: str, marcado: bool) -> None:
    ctrl = _obter_checkbox(class_name, texto)
    if not ctrl:
        _log(f"Checkbox '{texto}' ({class_name}) nao localizado/legivel.")
        _log_checkboxes_disponiveis()
        return
    estado = _check_state(ctrl)
    if estado is None:
        _log(f"Checkbox '{texto}' ({class_name}) sem estado legivel; tentando clique.")
        try:
            ctrl.click_input()
        except Exception:
            try:
                ctrl.click()
            except Exception:
                pass
        estado = _check_state(ctrl)
        if estado is None:
            _log(f"Checkbox '{texto}' ({class_name}) ainda sem estado legivel.")
            return
    if bool(estado) == marcado:
        _log(f"Checkbox '{texto}' ja esta {'marcado' if marcado else 'desmarcado'}.")
        return
    _log(f"Ajustando checkbox '{texto}' para {'marcado' if marcado else 'desmarcado'}.")
    try:
        ctrl.click_input()
    except Exception:
        ctrl.click()
    estado_final = _check_state(ctrl)
    if estado_final is None:
        _log(f"Checkbox '{texto}' ajuste sem confirmacao de estado.")
    elif bool(estado_final) != marcado:
        _log(f"Checkbox '{texto}' nao ficou no estado esperado.")


def _esperar_aviso_sucesso() -> bool:
    desk = Desktop(backend="win32")
    inicio = time.time()
    handles_anteriores: set[int] = set()

    try:
        avisos = findwindows.find_elements(
            backend="win32",
            class_name=AVISO_CLASS,
            title=AVISO_TITULO,
            top_level_only=True,
        )
        handles_anteriores = {int(aviso.handle) for aviso in avisos}
        if handles_anteriores:
            _log(f"Avisos ja abertos antes da confirmacao: {len(handles_anteriores)}")
    except Exception:
        handles_anteriores = set()

    while (time.time() - inicio) < WAIT_APOS_OK:
        try:
            avisos = findwindows.find_elements(
                backend="win32",
                class_name=AVISO_CLASS,
                title=AVISO_TITULO,
                top_level_only=True,
            )
            for aviso in avisos:
                handle = int(aviso.handle)
                if handle in handles_anteriores:
                    continue

                win = desk.window(handle=handle)
                if not win.exists(timeout=0.2):
                    continue

                _log("Aviso do Sistema final localizado.")
                try:
                    btn = win.child_window(class_name=AVISO_OK_CLASS)
                    if btn.exists(timeout=0.5):
                        _log("Botao OK do aviso localizado. Confirmando aviso.")
                        try:
                            btn.click_input()
                        except Exception:
                            btn.click()
                        return True
                except Exception:
                    pass

                _log("Botao OK do aviso nao localizado. Enviando ENTER.")
                try:
                    win.set_focus()
                except Exception:
                    pass
                pag.press("enter")
                return True
        except Exception:
            pass
        time.sleep(0.2)

    _log(f"Timeout aguardando o aviso final de sucesso ({WAIT_APOS_OK}s).")
    return False


# =====================
# Fluxo principal
# =====================

def main() -> int:
    empresa = _obter_empresa(sys.argv)
    if not empresa:
        _log("Empresa vazia. Encerrando.")
        return 1

    _click((401, 463), "clique inicial")
    sem_ok = _modo_sem_ok(sys.argv)

    _log(f"Iniciando: {NOME_SCRIPT}")
    _log(f"Empresa: {empresa}")
    _log(f"Diretorio base esperado: {_caminho_destino(empresa)}")
    _sleep(WAIT_START, "preparar a tela")

    _log("ALT + MPA")
    _alt_sequencia("mpa")

    _sleep(WAIT_CURTO)
    pag.hotkey("ctrl", "a")
    pag.write(str(ANO_PASSADO), interval=0.02)
    pag.press("enter", presses=3)
    pag.hotkey("ctrl", "a")
    pag.write("5", interval=0.02)
    pag.press("enter")
    _sleep(WAIT_CURTO)

    try:
        _garantir_checkbox(CHECK_GERAR_SALDO_ANTERIOR_CLASS, "Gerar Saldo Anterior", True)
        _garantir_checkbox(CHECK_EXCLUIR_DEMONSTRACOES_CLASS, "Excluir Demonstrações já Geradas nesse Exercício", True)
        _garantir_checkbox(CHECK_GERAR_CONTAS_SALDO_ZERO_CLASS, "Gerar Contas com Saldo Zero", False)
        _garantir_checkbox(CHECK_ABERTURA_EMPRESA_DIFERENTE_CLASS, "Abertura da Empresa Diferente de Janeiro no Ano Anterior", False)
        _garantir_checkbox(CHECK_MENSAL_CLASS, "Mensal", False)
        _garantir_checkbox(CHECK_TRIMESTRAL_CLASS, "Trimestral", False)
        _garantir_checkbox(CHECK_SEMESTRAL_CLASS, "Semestral", False)
        _garantir_checkbox(CHECK_ANUAL_CLASS, "Anual", True)
    except Exception as exc:
        _log(f"Falha ao ler/ajustar checkboxes: {exc}")

    if sem_ok:
        _log("Modo sem OK ativo. Encerrando antes do clique no OK.")
        return 0

    _sleep(DELAY_APOS_BIND, "delay apos bind")
    _click(BTN_OK, "ok")

    if not _esperar_aviso_sucesso():
        _log("Encerrando sem pos-processo porque o aviso final nao apareceu.")
        return 1

    _click(BTN_POS_PROCESSO_1, "pos-processo 1")
    _sleep(DELAY_ENTRE_CLIQUES, "delay entre cliques")
    _click(BTN_POS_PROCESSO_2, "pos-processo 2")

    _sleep(1, "finalizacao")
    _log("Finalizado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


