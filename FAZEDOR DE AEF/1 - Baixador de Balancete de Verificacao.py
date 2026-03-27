# 1 - Baixador de Balancete de Verificacao

# =========================
# Configuracoes
# =========================

import os
import sys
import time
from datetime import datetime
import tkinter as tk

try:
    import pyautogui
except Exception as exc:
    print("ERRO: dependencias nao encontradas.")
    print("Instale: pip install pyautogui")
    print(f"Detalhe: {exc}")
    raise

try:
    from pywinauto import Desktop
except Exception as exc:
    print("ERRO: pywinauto nao encontrado.")
    print("Instale: pip install pywinauto")
    print(f"Detalhe: {exc}")
    raise

try:
    import win32gui
except Exception as exc:
    print("ERRO: pywin32 nao encontrado.")
    print("Instale: pip install pywin32")
    print(f"Detalhe: {exc}")
    raise


TITULO_JANELA_CONTABILIDADE = "Contabilidade"
CLASS_NAME_CONTABILIDADE = "TfrmPrincipal"
USAR_CLASS_NAME_CONTABILIDADE = True

CAMINHO_EMPRESAS = (
    r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\FAZEDOR DE AEF\empresas.txt"
)

COORD_EMPRESA_SISTEMA = (1635, 58)
COORD_ABRIR_GERADOR = (509, 78)
COORD_BTN_SALVAR = (436, 85)
COORD_BTN_VOLTAR = (492, 85)
COORD_ESTABELECIMENTO = (1033, 508)

BASE_DIR_ARQUIVOS = (
    r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\FAZEDOR DE AEF\Arquivos"
)

QTDE_BACKSPACE = 4

DELAY_ENTRE_ACOES = 0.1
INTERVALO_CLIQUE = 0.0
PAUSA_APOS_ENTERS = 0.5
INTERVALO_DIGITACAO = 0.1
INTERVALO_TECLA = 0.1
PAUSA_PADRAO = 0.0
TIMEOUT_JANELA_IMPRESSAO = 300
TIMEOUT_SALVAR_COMO = 300
TIMEOUT_FECHAR_IMPRESSAO = 200
TIMEOUT_DETECCAO_ARQUIVO = 600
PAUSAR_APOS_DETECTAR_IMPRESSAO = False
ENVIAR_ENTERS_ENTRE_CLIQUES = True
PAUSAR_APOS_DIGITAR_CAMINHO = False
PULAR_FOCO_NAS_PROXIMAS_EMPRESAS = True
GARANTIR_CHECKBOX_DESMARCADA = True
CHECKBOX_NOME = "TCRDCheckBox1"
CHECKBOX_CLASSE = "TCRDCheckBox"
CHECKBOX_TEXTO = "tirar D e C"

# =========================
# Utilitarios
# =========================


def carregar_empresas(caminho: str) -> list[str]:
    try:
        with open(caminho, "r", encoding="utf-8") as arquivo:
            linhas = [_normalizar_empresa(linha) for linha in arquivo.readlines()]
    except FileNotFoundError:
        print(f"ERRO: arquivo de empresas nao encontrado: {caminho}")
        sys.exit(1)

    empresas = [linha for linha in linhas if linha]
    if not empresas:
        print("ERRO: lista de empresas vazia.")
        sys.exit(1)

    return empresas


def _normalizar_empresa(texto: str) -> str:
    return texto.strip().lstrip("\ufeff")


def _separar_empresa_estabelecimento(empresa: str) -> tuple[str, str | None]:
    """
    Regras:
    - "143-1" => empresa sistema "143" e estabelecimento "1"
    - "143-2" => empresa sistema "143" e estabelecimento "2"
    - demais => usa o valor original e sem estabelecimento adicional
    """
    valor = _normalizar_empresa(empresa)
    if "-" in valor:
        base, sufixo = valor.split("-", 1)
        base = base.strip()
        sufixo = sufixo.strip()
        if base == "143" and sufixo in ("1", "2"):
            return base, sufixo
    return valor, None


def calcular_periodo() -> tuple[str, str]:
    """
    Regra:
    - Janeiro: 01/AA-1 ate 12/AA-1
    - Fev..Dez: 01/AA ate MM-1/AA
      Ex.: fev/2026 => 01/26 ate 01/26
            mar/2026 => 01/26 ate 02/26
    """
    agora = datetime.now()
    ano_atual = agora.year
    mes_atual = agora.month

    if mes_atual == 1:
        ano_base = ano_atual - 1
        data_inicio = f"01/{str(ano_base)[-2:]}"
        data_fim = f"12/{str(ano_base)[-2:]}"
        return data_inicio, data_fim

    mes_fim = mes_atual - 1
    ano_curto = str(ano_atual)[-2:]
    data_inicio = f"01/{ano_curto}"
    data_fim = f"{mes_fim:02d}/{ano_curto}"
    return data_inicio, data_fim


def set_clipboard(texto: str) -> None:
    root = tk.Tk()
    root.withdraw()
    root.clipboard_clear()
    root.clipboard_append(texto)
    root.update()
    root.destroy()


def obter_janela_principal():
    try:
        for w in Desktop(backend="win32").windows():
            try:
                titulo = w.window_text() or ""
                if TITULO_JANELA_CONTABILIDADE.lower() not in titulo.lower():
                    continue
                if USAR_CLASS_NAME_CONTABILIDADE:
                    if w.class_name() == CLASS_NAME_CONTABILIDADE:
                        return w
                else:
                    return w
            except Exception:
                continue
    except Exception:
        pass

    if USAR_CLASS_NAME_CONTABILIDADE:
        try:
            for w in Desktop(backend="win32").windows():
                try:
                    titulo = w.window_text() or ""
                    if TITULO_JANELA_CONTABILIDADE.lower() in titulo.lower():
                        return w
                except Exception:
                    continue
        except Exception:
            pass
    return None


def snapshot_children(hwnd: int) -> set[int]:
    handles: set[int] = set()
    try:
        def _cb(h, lparam):
            handles.add(h)
            return True
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass
    return handles


def focar_janela_principal() -> int:
    janela = obter_janela_principal()
    if not janela:
        print(f"ERRO: janela nao encontrada: '{TITULO_JANELA_CONTABILIDADE}'")
        sys.exit(1)

    hwnd = janela.handle
    try:
        janela.restore()
        janela.set_focus()
    except Exception:
        pass

    try:
        win32gui.ShowWindow(hwnd, 3)
        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
        win32gui.SetActiveWindow(hwnd)
    except Exception:
        pass

    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        x = left + 60
        y = top + 10
        pyautogui.click(x, y)
    except Exception:
        pass

    time.sleep(PAUSA_PADRAO)
    return hwnd


def aguardar_janela_impressao(
    timeout_s: int, baseline: set[int], hwnd: int | None
) -> set[int]:
    inicio = time.time()
    while time.time() - inicio < timeout_s:
        try:
            if hwnd:
                atuais = snapshot_children(hwnd)
                novos = atuais - baseline
                if novos:
                    print(f"Janela de impressao detectada (novo child window: {len(novos)}).")
                    return novos
        except Exception:
            pass
        time.sleep(1)

    print("ERRO: timeout aguardando janela de impressao.")
    sys.exit(1)


def aguardar_fechamento_impressao(
    timeout_s: int, hwnd: int | None, handles_impressao: set[int]
) -> None:
    if not handles_impressao:
        return

    inicio = time.time()
    while time.time() - inicio < timeout_s:
        try:
            if hwnd:
                atuais = snapshot_children(hwnd)
                ainda_abertos = handles_impressao.intersection(atuais)
                if not ainda_abertos:
                    return
        except Exception:
            pass
        time.sleep(0.5)

    print("AVISO: janela de impressao ainda aberta.")
    input("Feche a janela e pressione ENTER para continuar.")


def obter_checkbox(janela) -> object | None:
    try:
        chk = janela.child_window(best_match=CHECKBOX_NOME)
        if chk.exists(timeout=0.2):
            return chk.wrapper_object()
    except Exception:
        pass

    try:
        candidatos = []
        for c in janela.children():
            try:
                if c.class_name() != CHECKBOX_CLASSE:
                    continue
                if CHECKBOX_TEXTO and CHECKBOX_TEXTO.lower() not in (c.window_text() or "").lower():
                    continue
                candidatos.append(c)
            except Exception:
                continue

        if len(candidatos) == 1:
            return candidatos[0]
        if len(candidatos) > 1:
            print("ERRO: mais de um checkbox encontrado. Informe um texto para identificar.")
            sys.exit(1)
    except Exception:
        pass

    return None


def garantir_checkbox_desmarcada(hwnd: int) -> None:
    if not GARANTIR_CHECKBOX_DESMARCADA:
        return

    try:
        janela = Desktop(backend="win32").window(handle=hwnd)
    except Exception:
        print("ERRO: nao foi possivel acessar a janela principal.")
        sys.exit(1)

    chk = obter_checkbox(janela)
    if not chk:
        print("ERRO: checkbox nao encontrada.")
        sys.exit(1)

    try:
        estado = chk.get_check_state()
        if estado != 0:
            chk.uncheck()
        return
    except Exception:
        pass

    try:
        chk.click_input()
    except Exception:
        print("ERRO: nao foi possivel desmarcar a checkbox.")
        sys.exit(1)


def aguardar_salvar_como(timeout_s: int):
    t0 = time.time()
    while time.time() - t0 < timeout_s:
        try:
            dlg = Desktop(backend="win32").window(
                class_name="#32770", title_re=".*Salvar como.*|.*Save As.*"
            )
            if dlg.exists(timeout=0.2):
                return dlg
        except Exception:
            pass
        time.sleep(0.3)
    print("ERRO: timeout aguardando janela de salvar.")
    sys.exit(1)


def focar_nome_arquivo(dlg) -> None:
    try:
        dlg.set_focus()
    except Exception:
        pass
    try:
        edit = dlg.child_window(class_name="Edit", found_index=0)
        edit.set_focus()
        edit.click_input()
        return
    except Exception:
        pass
    pyautogui.hotkey("alt", "n")


def definir_nome_arquivo(dlg, nome_arquivo: str) -> bool:
    try:
        edit = dlg.child_window(class_name="Edit", found_index=0).wrapper_object()
        edit.set_edit_text(nome_arquivo)
        return True
    except Exception:
        return False


def salvar_arquivo(caminho_completo: str) -> None:
    dlg = aguardar_salvar_como(TIMEOUT_SALVAR_COMO)

    focar_nome_arquivo(dlg)
    time.sleep(0.2)
    caminho = caminho_completo.replace("\\\\", "\\")
    if not definir_nome_arquivo(dlg, caminho):
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("backspace")
        pyautogui.write(caminho, interval=INTERVALO_DIGITACAO)
    time.sleep(0.3)

    if PAUSAR_APOS_DIGITAR_CAMINHO:
        input("Caminho digitado. Verifique e pressione ENTER para continuar.")
        return

    pyautogui.press("enter")
    time.sleep(0.5)

    encontrado, caminho_encontrado = aguardar_arquivo(caminho_completo, TIMEOUT_DETECCAO_ARQUIVO)
    if encontrado:
        if caminho_encontrado != caminho_completo:
            print(f"AVISO: arquivo localizado com nome diferente: {caminho_encontrado}")
        return

    print(f"ERRO: arquivo nao encontrado apos salvar: {caminho_completo}")
    input("Pressione ENTER para continuar.")


def aguardar_arquivo(caminho_completo: str, timeout_s: int) -> tuple[bool, str]:
    caminho_completo = os.path.normpath(caminho_completo)
    pasta = os.path.dirname(caminho_completo)
    base_path = os.path.splitext(caminho_completo)[0]
    base_nome = os.path.basename(base_path).lower()
    caminho_xls = base_path + ".xls"
    caminho_xlsx = base_path + ".xlsx"
    caminho_xls_xlsx = base_path + ".xls.xlsx"
    caminho_xlsx_xls = base_path + ".xlsx.xls"

    def _base_sem_ext(nome_arquivo: str) -> str:
        nome = nome_arquivo.lower()
        for _ in range(2):
            root, ext = os.path.splitext(nome)
            if ext in [".xls", ".xlsx"]:
                nome = root
            else:
                break
        return nome

    t0 = time.time()
    while time.time() - t0 < timeout_s:
        if os.path.isfile(caminho_completo):
            return True, caminho_completo
        if os.path.isfile(caminho_xls):
            return True, caminho_xls
        if os.path.isfile(caminho_xlsx):
            return True, caminho_xlsx
        if os.path.isfile(caminho_xls_xlsx):
            return True, caminho_xls_xlsx
        if os.path.isfile(caminho_xlsx_xls):
            return True, caminho_xlsx_xls

        try:
            if os.path.isdir(pasta):
                for nome in os.listdir(pasta):
                    base_atual = _base_sem_ext(nome)
                    if base_atual == base_nome:
                        return True, os.path.join(pasta, nome)
        except Exception:
            pass

        time.sleep(1)

    return False, caminho_completo


def executar_fluxo(empresa: str, primeira_empresa: bool) -> None:
    data_inicio, data_fim = calcular_periodo()
    empresa_id = empresa.strip()
    empresa_sistema, estab = _separar_empresa_estabelecimento(empresa_id)
    destino = os.path.join(BASE_DIR_ARQUIVOS, empresa_id)
    os.makedirs(destino, exist_ok=True)
    nome_arquivo = f"Balancete_{empresa_id}.xls"
    caminho_arquivo = os.path.join(destino, nome_arquivo)

    if not (PULAR_FOCO_NAS_PROXIMAS_EMPRESAS and not primeira_empresa):
        hwnd_principal = focar_janela_principal()
    else:
        hwnd_principal = obter_janela_principal()
        if not hwnd_principal:
            print(f"ERRO: janela nao encontrada: '{TITULO_JANELA_CONTABILIDADE}'")
            sys.exit(1)
        hwnd_principal = hwnd_principal.handle

    pyautogui.click(*COORD_EMPRESA_SISTEMA)
    time.sleep(INTERVALO_CLIQUE)
    pyautogui.press("backspace", presses=QTDE_BACKSPACE, interval=INTERVALO_TECLA)
    pyautogui.write(empresa_sistema, interval=INTERVALO_DIGITACAO)

    if ENVIAR_ENTERS_ENTRE_CLIQUES:
        pyautogui.press("enter")
        pyautogui.press("enter")
        time.sleep(PAUSA_APOS_ENTERS)

    pyautogui.click(*COORD_ABRIR_GERADOR)
    time.sleep(INTERVALO_CLIQUE)

    # Regra especifica 143-1 / 143-2:
    # apos abrir gerador, antes do Enter, digita o estabelecimento.
    if estab is not None:
        pyautogui.write(estab, interval=INTERVALO_DIGITACAO)

    pyautogui.press("enter")

    pyautogui.write(data_inicio, interval=INTERVALO_DIGITACAO)
    pyautogui.press("enter")
    pyautogui.write(data_fim, interval=INTERVALO_DIGITACAO)
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.write("6", interval=INTERVALO_DIGITACAO)
    if hwnd_principal:
        garantir_checkbox_desmarcada(hwnd_principal)
    # Antes de imprimir, clica no campo de estabelecimento conforme solicitado.
    pyautogui.click(*COORD_ESTABELECIMENTO)
    time.sleep(INTERVALO_CLIQUE)
    handles_antes = snapshot_children(hwnd_principal) if hwnd_principal else set()
    pyautogui.hotkey("alt", "i")

    handles_impressao = aguardar_janela_impressao(
        TIMEOUT_JANELA_IMPRESSAO, handles_antes, hwnd_principal
    )

    pyautogui.click(*COORD_BTN_SALVAR)
    time.sleep(INTERVALO_CLIQUE)
    salvar_arquivo(caminho_arquivo)

    if PAUSAR_APOS_DETECTAR_IMPRESSAO:
        input("Pressione ENTER para continuar.")

    pyautogui.click(*COORD_BTN_VOLTAR)
    time.sleep(INTERVALO_CLIQUE)
    aguardar_fechamento_impressao(
        TIMEOUT_FECHAR_IMPRESSAO, hwnd_principal, handles_impressao
    )


def main() -> None:
    pyautogui.PAUSE = DELAY_ENTRE_ACOES
    pyautogui.FAILSAFE = True

    empresas = carregar_empresas(CAMINHO_EMPRESAS)

    for idx, empresa in enumerate(empresas, start=1):
        print(f"Processando empresa: {empresa}")
        executar_fluxo(empresa, primeira_empresa=(idx == 1))


if __name__ == "__main__":
    main()
