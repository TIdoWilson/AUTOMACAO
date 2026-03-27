# -*- coding: utf-8 -*-
"""
Relatório de Notas Explicativas ECD
"""

from __future__ import annotations

import datetime as _dt
import os
import sys
import time
import tkinter as tk

try:
    import pyautogui as pag
except Exception as exc:
    raise SystemExit("pyautogui é obrigatório para a automação de UI") from exc

try:
    from pywinauto import Desktop
except Exception as exc:
    raise SystemExit("pywinauto é obrigatório para focar o campo de salvar") from exc

try:
    from pywinauto import findwindows
except Exception as exc:
    raise SystemExit("pywinauto é obrigatório para detectar o relatório") from exc

try:
    import win32com.client as win32
except Exception as exc:
    raise SystemExit("pywin32 é obrigatório para formatar o RTF") from exc

try:
    import win32gui
except Exception as exc:
    raise SystemExit("pywin32 é obrigatório para detectar janela do relatório") from exc
# =====================
# Configuracoes
# =====================
NOME_SCRIPT = "1 - Relatório de Notas Explicativas ECD"
BASE_DIR = r"W:\\SPEDs\\ECD\\2025"
ANO_PASSADO = _dt.date.today().year - 1
LOG_PATH = os.path.join(os.path.dirname(__file__), "log_notas_explicativas_ecd.txt")

# Coordenadas
BTN_LISTAR_ASSINATURA = (792, 570)
BTN_LISTAR_CNPJ = (791, 585)
BTN_IMPRIMIR = (1161, 464)
BTN_SALVAR = (405, 87)
BTN_VOLTAR_1 = (495, 81)
BTN_VOLTAR_2 = (1214, 438)

# Tempos
WAIT_START = 2
WAIT_APOS_IMPRIMIR = 1
WAIT_APOS_SALVAR = 1.5
WAIT_APOS_VOLTAR = 1
WAIT_CURTO = 0.5
WAIT_APOS_ALT_RNX = 1
WAIT_APOS_ENTER_ANO = 0.5
TIMEOUT_RELATORIO = 3600

MAX_TENTATIVAS = 2
NOME_ARQUIVO = "Nota_explicativa.rtf"

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


def _set_clipboard(texto: str) -> None:
    root = tk.Tk()
    root.withdraw()
    root.clipboard_clear()
    root.clipboard_append(texto)
    root.update()
    root.destroy()

def _formatar_rtf(caminho_arquivo: str) -> None:
    import re
    wd_field_page = 33
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(caminho_arquivo, ReadOnly=False)
        # Fonte 7 no documento inteiro
        doc.Range().Font.Size = 7

        def _substituir_matches_por_espacos(texto: str, pattern: str) -> list[tuple[int, int]]:
            return [(m.start(), m.end()) for m in re.finditer(pattern, texto)]

        def _aplicar_substituicoes(doc_range, posicoes: list[tuple[int, int]]) -> None:
            for ini, fim in reversed(posicoes):
                rng = doc.Range(ini, fim)
                rng.Text = " " * max(0, fim - ini)

        # Remove campos de pagina (se existirem)
        try:
            for fld in list(doc.Fields):
                if fld.Type == wd_field_page:
                    fld.Delete()
        except Exception:
            pass

        # Remove "Pagina/Página" + numero mantendo layout (espacos)
        try:
            texto = doc.Content.Text
            pos = []
            pos += _substituir_matches_por_espacos(texto, r"(?i)p[áa]gina:\s*\d+")
            pos += _substituir_matches_por_espacos(texto, r"(?i)p[áa]gina\s+\d+")
            pos += _substituir_matches_por_espacos(texto, r"(?i)p[áa]gina:")
            _aplicar_substituicoes(doc.Content, pos)
        except Exception:
            pass

        # Limpar em todas as histórias do Word (main, headers, footers, textboxes, etc.)
        try:
            for i in range(1, doc.StoryRanges.Count + 1):
                story = doc.StoryRanges(i)
                while story is not None:
                    try:
                        texto = story.Text
                        pos = []
                        pos += _substituir_matches_por_espacos(texto, r"(?i)p[áa]gina:\s*\d+")
                        pos += _substituir_matches_por_espacos(texto, r"(?i)p[áa]gina\s+\d+")
                        pos += _substituir_matches_por_espacos(texto, r"(?i)p[áa]gina:")
                        _aplicar_substituicoes(story, pos)
                    except Exception:
                        pass
                    try:
                        story = story.NextStoryRange
                    except Exception:
                        break
        except Exception:
            pass

        # Remover marcacao de paginas em cabecalhos/rodapes (reforco)
        # Shapes no documento
        try:
            for shp in doc.Shapes:
                if shp.TextFrame.HasText:
                    tr = shp.TextFrame.TextRange
                    texto = tr.Text
                    pos = []
                    pos += _substituir_matches_por_espacos(texto, r"(?i)p[áa]gina:\s*\d+")
                    pos += _substituir_matches_por_espacos(texto, r"(?i)p[áa]gina\s+\d+")
                    pos += _substituir_matches_por_espacos(texto, r"(?i)p[áa]gina:")
                    for ini, fim in reversed(pos):
                        sub = " " * max(0, fim - ini)
                        tr.Characters(ini + 1, fim - ini).Text = sub
        except Exception:
            pass
        doc.Save()
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()

def _obter_handle_principal() -> int | None:
    try:
        for w in Desktop(backend="win32").windows():
            try:
                if w.class_name() == "TfrmPrincipal" and "Contabilidade" in (w.window_text() or ""):
                    return w.handle
            except Exception:
                continue
    except Exception:
        return None
    return None


def _snapshot_children(hwnd: int) -> set[int]:
    handles: set[int] = set()
    try:
        def _cb(h, lparam):
            handles.add(h)
            return True
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass
    return handles


def _esperar_relatorio(timeout: int, baseline: set[int]) -> bool:
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            hwnd = _obter_handle_principal()
            if hwnd:
                atuais = _snapshot_children(hwnd)
                novos = atuais - baseline
                if novos:
                    _log(f"Relatorio localizado (novo child window: {len(novos)}).")
                    return True
        except Exception:
            pass
        time.sleep(3)
    _log(f"ERRO: timeout aguardando relatorio ({timeout}s).")
    return False

def _focar_campo_salvar_como() -> None:
    try:
        dlg = Desktop(backend="win32").window(class_name="#32770", title_re=".*Salvar como.*|.*Save As.*")
        if dlg.exists(timeout=2):
            edit = dlg.child_window(class_name="Edit")
            edit.set_focus()
            edit.click_input()
    except Exception:
        pass


def _focar_nome_arquivo() -> None:
    try:
        dlg = Desktop(backend="win32").window(class_name="#32770", title_re=".*Salvar como.*|.*Save As.*")
        if dlg.exists(timeout=2):
            edit = dlg.child_window(class_name="Edit", found_index=0)
            edit.set_focus()
            edit.click_input()
            return
    except Exception:
        pass
    pag.hotkey("alt", "n")


def _definir_nome_arquivo(nome_arquivo: str) -> bool:
    try:
        dlg = Desktop(backend="win32").window(class_name="#32770", title_re=".*Salvar como.*|.*Save As.*")
        if dlg.exists(timeout=2):
            edit = dlg.child_window(class_name="Edit", found_index=0).wrapper_object()
            edit.set_edit_text(nome_arquivo)
            return True
    except Exception:
        return False
    return False


def _focar_barra_endereco() -> None:
    try:
        dlg = Desktop(backend="win32").window(class_name="#32770", title_re=".*Salvar como.*|.*Save As.*")
        if dlg.exists(timeout=2):
            toolbar = dlg.child_window(class_name="ToolbarWindow32")
            toolbar.set_focus()
            toolbar.click_input()
    except Exception:
        pass


def _salvar_arquivo(caminho_completo: str) -> None:
    os.makedirs(os.path.dirname(caminho_completo), exist_ok=True)
    pasta = os.path.dirname(caminho_completo).replace("\\\\", "\\")
    nome_arquivo = os.path.basename(caminho_completo)
    # Barra de endereco
    _focar_barra_endereco()
    pag.hotkey("ctrl", "l")
    pag.hotkey("ctrl", "a")
    # Forca uso do drive W:
    _set_clipboard("W:")
    pag.hotkey("ctrl", "v")
    time.sleep(0.05)
    pag.press("enter")
    time.sleep(0.1)
    pag.hotkey("ctrl", "l")
    pag.hotkey("ctrl", "a")
    _set_clipboard(pasta)
    pag.hotkey("ctrl", "v")
    time.sleep(0.05)
    pag.press("enter")
    time.sleep(0.1)
    pag.press("enter")
    time.sleep(0.5)
    # Campo nome do arquivo
    _focar_nome_arquivo()
    time.sleep(0.2)
    if not _definir_nome_arquivo(nome_arquivo):
        pag.hotkey("ctrl", "a")
        pag.press("backspace")
        _set_clipboard(nome_arquivo)
        pag.hotkey("ctrl", "v")
    time.sleep(1)
    pag.press("enter")
    time.sleep(0.5)


# =====================
# Fluxo principal
# =====================

def main() -> int:
    empresa = _obter_empresa(sys.argv)
    if not empresa:
        _log("Empresa vazia. Encerrando.")
        return 1

    _click((401, 463), "clique inicial")

    destino = _caminho_destino(empresa)
    os.makedirs(destino, exist_ok=True)
    caminho_arquivo = os.path.join(destino, NOME_ARQUIVO)

    _log(f"Iniciando: {NOME_SCRIPT}")
    _log(f"Empresa: {empresa}")
    _log(f"Diretorio destino: {destino}")
    _sleep(WAIT_START, "preparar a tela")

    for tentativa in range(1, MAX_TENTATIVAS + 1):
        _log(f"Tentativa {tentativa}/{MAX_TENTATIVAS}")

        _log("ALT + RNX")
        _alt_sequencia("rnx")
        _sleep(WAIT_APOS_ALT_RNX, "aguardar bind apos ALT+RNX")

        pag.press("enter")
        _sleep(WAIT_APOS_ENTER_ANO, "aguardar campo do ano")
        pag.write(str(ANO_PASSADO), interval=0.02)
        pag.press("enter")
        _sleep(WAIT_CURTO)

        _click(BTN_LISTAR_ASSINATURA, "listar dados para assinatura")
        _click(BTN_LISTAR_CNPJ, "listar CNPJ")
        hwnd = _obter_handle_principal()
        baseline = _snapshot_children(hwnd) if hwnd else set()
        _click(BTN_IMPRIMIR, "imprimir")
        _esperar_relatorio(TIMEOUT_RELATORIO, baseline)

        _click(BTN_SALVAR, "salvar")
        _sleep(WAIT_APOS_SALVAR, "janela de salvar")
        _salvar_arquivo(caminho_arquivo)
        _sleep(WAIT_CURTO)

        if os.path.isfile(caminho_arquivo):
            _log(f"Arquivo salvo: {caminho_arquivo}")
            _formatar_rtf(caminho_arquivo)
            sucesso = True
        else:
            _log("Arquivo nao apareceu no destino; continuando para retorno e nova tentativa")
            sucesso = False

        _click(BTN_VOLTAR_1, "voltar 1")
        _click(BTN_VOLTAR_2, "voltar 2")
        _sleep(WAIT_APOS_VOLTAR, "pos-voltar")

        if sucesso:
            _log("Finalizado com sucesso.")
            return 0

    _log("Finalizado sem sucesso. Reexecutar manualmente.")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())


