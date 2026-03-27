# -*- coding: utf-8 -*-
"""
DMED - Parte 1 (resumo extremo)
"""

from __future__ import annotations

import os
import time
from datetime import datetime
import ctypes

import pyautogui

try:
    from pywinauto import Desktop
except Exception as exc:
    Desktop = None
    print(f"Falha ao importar pywinauto: {exc}")

# ===================== CONFIGURACAO =====================
TITULO_JANELA = "Fiscal"
COORD_EMPRESA_SISTEMA = (1637, 58)
COORD_FOCO = (1600, 200)
COORD_CHECKBOX_FALLBACK = (815, 622)
COORD_SALVAR_FLUXO_FOCO = (671, 326)
COORD_SALVAR_FLUXO_BOTAO_SALVAR = (402, 83)
COORD_SALVAR_FLUXO_VOLTAR = (467, 85)
COORD_SALVAR_FLUXO_CONFIRMAR = (1745, 1007)
BACKSPACES_EMPRESA = 5
DELAY_POS_ALT_MGM = 0.2
TIMEOUT_AVISO_1 = 10
TIMEOUT_AVISO_2 = 5
DELAY_ENTRE_TEXTO_E_ENTER = 0.1
DELAY_APOS_ALT_G = 0.2
DELAY_APOS_VOLTAR = 0.1

DIR_BASE_TEMPLATE = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\DMED Gerar-Organizar-Formatar\arquivos base"
LISTA_TEMPLATE = r"W:\DECLARAÇÕES\DMED\DMED {ano}\LISTA.xlsx"
# ========================================================


def focar_janela_fiscal():
    if Desktop is None:
        print("pywinauto nao disponivel. Foque manualmente a janela 'Fiscal'.")
        return None

    try:
        dlg = Desktop(backend="uia").window(title_re=f".*{TITULO_JANELA}.*")
        if not dlg.exists(timeout=3):
            print("Janela 'Fiscal' nao encontrada. Foque manualmente.")
            return None
        dlg.set_focus()
        try:
            dlg.maximize()
        except Exception:
            pass
        return dlg
    except Exception as exc:
        print(f"Falha ao focar a janela 'Fiscal': {exc}")
        return None


def clicar(x: int, y: int) -> None:
    pyautogui.click(x, y)


def escrever(texto: str, intervalo: float = 0.02) -> None:
    pyautogui.write(texto, interval=intervalo)


def escrever_unicode(texto: str, intervalo: float = 0.01) -> None:
    KEYEVENTF_UNICODE = 0x0004
    KEYEVENTF_KEYUP = 0x0002

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", ctypes.c_ushort),
            ("wScan", ctypes.c_ushort),
            ("dwFlags", ctypes.c_ulong),
            ("time", ctypes.c_ulong),
            ("dwExtraInfo", ctypes.c_void_p),
        ]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", ctypes.c_ulong), ("ki", KEYBDINPUT)]

    def _send_char(ch: str) -> None:
        code = ord(ch)
        inp_down = INPUT(1, KEYBDINPUT(0, code, KEYEVENTF_UNICODE, 0, None))
        inp_up = INPUT(1, KEYBDINPUT(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, None))
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp_down), ctypes.sizeof(inp_down))
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp_up), ctypes.sizeof(inp_up))

    for ch in texto:
        _send_char(ch)
        if intervalo:
            time.sleep(intervalo)


def _set_clipboard_text(texto: str) -> None:
    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002

    buf = ctypes.create_unicode_buffer(texto)
    tamanho = ctypes.sizeof(buf)

    if not ctypes.windll.user32.OpenClipboard(None):
        raise RuntimeError("Nao foi possivel abrir a area de transferencia.")
    try:
        ctypes.windll.user32.EmptyClipboard()
        hglobal = ctypes.windll.kernel32.GlobalAlloc(GMEM_MOVEABLE, tamanho)
        if not hglobal:
            raise RuntimeError("Falha ao alocar memoria para a area de transferencia.")
        lp = ctypes.windll.kernel32.GlobalLock(hglobal)
        if not lp:
            raise RuntimeError("Falha ao travar memoria da area de transferencia.")
        ctypes.memmove(lp, buf, tamanho)
        ctypes.windll.kernel32.GlobalUnlock(hglobal)
        ctypes.windll.user32.SetClipboardData(CF_UNICODETEXT, hglobal)
    finally:
        ctypes.windll.user32.CloseClipboard()


def colar_focado(texto: str) -> bool:
    try:
        _set_clipboard_text(texto)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.hotkey("ctrl", "v")
        return True
    except Exception:
        return False


def mandar_backspaces(qtd: int) -> None:
    for _ in range(qtd):
        pyautogui.press("backspace")


def enter_com_delay(qtd: int = 1) -> None:
    for _ in range(qtd):
        pyautogui.press("enter")
        time.sleep(DELAY_ENTRE_TEXTO_E_ENTER)


def delay_pos_alt_g() -> None:
    time.sleep(DELAY_APOS_ALT_G)


def alt_mgm() -> None:
    pyautogui.keyDown("alt")
    pyautogui.press("m")
    pyautogui.press("g")
    pyautogui.press("m")
    pyautogui.keyUp("alt")


def garantir_checkbox(dlg) -> None:
    clicar(*COORD_CHECKBOX_FALLBACK)


def _achar_aviso_sistema(timeout: float):
    if Desktop is None:
        time.sleep(timeout)
        return None
    titulo_regex = ".*Aviso.*Sistema.*"
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            # Tenta por UIA direto
            dlg = Desktop(backend="uia").window(title_re=titulo_regex)
            if dlg.exists(timeout=0.2):
                return dlg
        except Exception:
            pass

        try:
            # Varre janelas visiveis (UIA)
            for w in Desktop(backend="uia").windows():
                try:
                    if not w.is_visible():
                        continue
                    titulo = w.window_text() or ""
                    if "Aviso" in titulo and "Sistema" in titulo:
                        return w
                except Exception:
                    continue
        except Exception:
            pass

        try:
            # Fallback win32
            dlg = Desktop(backend="win32").window(title_re=titulo_regex)
            if dlg.exists(timeout=0.2):
                return dlg
        except Exception:
            pass

        time.sleep(0.1)
    return None


def _confirmar_aviso(dlg) -> None:
    if dlg is None:
        return
    try:
        btn_ok = dlg.child_window(title="OK", control_type="Button")
        if btn_ok.exists(timeout=1):
            btn_ok.click_input()
            return
    except Exception:
        pass
    pyautogui.press("enter")


def _esperar_e_confirmar_aviso(timeout: float) -> bool:
    dlg = _achar_aviso_sistema(timeout)
    if dlg is None:
        return False
    _confirmar_aviso(dlg)
    return True


def _fluxo_salvar_por_coordenadas(caminho: str, nome_arquivo: str) -> None:
    clicar(*COORD_SALVAR_FLUXO_FOCO)
    clicar(*COORD_SALVAR_FLUXO_BOTAO_SALVAR)
    time.sleep(3.0)
    caminho_completo = os.path.join(caminho, nome_arquivo)
    escrever(caminho_completo)
    enter_com_delay()
    clicar(*COORD_SALVAR_FLUXO_CONFIRMAR)
    clicar(*COORD_SALVAR_FLUXO_VOLTAR)


def registrar_movimento(dir_dest: str, empresa: str) -> None:
    pasta_pai = dir_dest.rstrip("\\/").rsplit("\\", 1)[0]
    caminho_txt = f"{pasta_pai}\\DMED Movimentos.txt"
    try:
        with open(caminho_txt, "a", encoding="utf-8") as f_out:
            f_out.write(f"{empresa}\n")
    except Exception as exc:
        print(f"Falha ao registrar movimento em {caminho_txt}: {exc}")


def carregar_empresas_lista(caminho_lista: str) -> list[str]:
    try:
        from openpyxl import load_workbook
    except Exception:
        print("openpyxl nao encontrado. Instale para ler a LISTA.xlsx.")
        return []

    if not os.path.exists(caminho_lista):
        print(f"Lista nao encontrada: {caminho_lista}")
        return []

    wb = load_workbook(caminho_lista, data_only=True)
    ws = wb.active
    empresas: list[str] = []

    for row in ws.iter_rows(min_col=3, max_col=3, values_only=True):
        val = row[0]
        if val is None:
            continue
        texto = str(val).strip()
        if not texto:
            continue
        empresas.append(texto)

    return empresas


def processar_empresa(empresa: str, dir_dest: str, ano_atual: int, dlg) -> None:
    nome_arquivo = f"{empresa}.txt"
    nome_arquivo_slk = f"{empresa}.slk"

    clicar(*COORD_EMPRESA_SISTEMA)
    mandar_backspaces(BACKSPACES_EMPRESA)
    escrever(empresa)
    enter_com_delay(2)
    clicar(*COORD_FOCO)

    alt_mgm()

    time.sleep(DELAY_POS_ALT_MGM)

    escrever(str(ano_atual))
    enter_com_delay()

    escrever(str(ano_atual - 1))
    enter_com_delay()

    enter_com_delay(3)

    escrever(dir_dest)
    enter_com_delay()

    escrever(nome_arquivo)
    enter_com_delay()

    garantir_checkbox(dlg)

    pyautogui.hotkey("alt", "g")
    delay_pos_alt_g()

    _esperar_e_confirmar_aviso(TIMEOUT_AVISO_1)
    delay_pos_alt_g()
    if _esperar_e_confirmar_aviso(TIMEOUT_AVISO_2):
        delay_pos_alt_g()
        return
    delay_pos_alt_g()
    _fluxo_salvar_por_coordenadas(dir_dest, nome_arquivo_slk)
    time.sleep(DELAY_APOS_VOLTAR)


def main() -> None:
    import argparse

    p = argparse.ArgumentParser()
    p.add_argument("--empresa", "-e", default="")
    args = p.parse_args()

    ano_atual = datetime.now().year

    dir_dest = DIR_BASE_TEMPLATE.format(ano=ano_atual)
    caminho_lista = LISTA_TEMPLATE.format(ano=ano_atual)

    dlg = focar_janela_fiscal()
    time.sleep(0.3)

    empresas: list[str] = []
    if args.empresa:
        empresas = [str(args.empresa).strip()]
    else:
        empresas = carregar_empresas_lista(caminho_lista)

    for empresa in empresas:
        if not empresa:
            continue
        caminho_txt = os.path.join(dir_dest, f"{empresa}.txt")
        if os.path.exists(caminho_txt):
            print(f"Ja existe TXT para {empresa}. Pulando.")
            continue
        processar_empresa(empresa, dir_dest, ano_atual, dlg)


if __name__ == "__main__":
    main()
