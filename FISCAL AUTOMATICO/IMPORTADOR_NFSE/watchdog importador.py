# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess
import unicodedata
from collections import deque

import tkinter as tk
from tkinter import filedialog

import pyautogui as pag
import uiautomation as uia


NOME_ALERTA = "aviso do sistema"
GERENCIADOR_TITLE = "gerenciador de sistemas"
FISCAL_TITLE = "fiscal"
MAX_RESTARTS = 40
RESTART_DELAY = 2.0
CHECKPOINT_NOME_ARQUIVO = "watchdog_importador_checkpoint.txt"
CODIGO_RECUPERADO_CRASH_FISCAL = -998

# Coordenada validada manualmente pelo usuario para o botao "Fiscal" no Gerenciador.
BOTAO_FISCAL_ABS_X = 954
BOTAO_FISCAL_ABS_Y = 179
# Fallback proporcional ao tamanho da janela do Gerenciador.
BOTAO_FISCAL_RATIO_X = 0.497
BOTAO_FISCAL_RATIO_Y = 0.175


def _agora() -> str:
    return time.strftime("%H:%M:%S")


def _log(msg: str) -> None:
    print(f"[{_agora()}][WATCHDOG] {msg}", flush=True)


def _normalizar_txt(txt: str) -> str:
    txt = txt or ""
    txt = "".join(
        ch for ch in unicodedata.normalize("NFD", txt)
        if unicodedata.category(ch) != "Mn"
    )
    return " ".join(txt.lower().split())


def escolher_script_importador() -> str:
    diretorio_padrao = os.path.dirname(os.path.abspath(__file__))
    arquivo_padrao = "importador final com fator 2.py"
    caminho_padrao = os.path.join(diretorio_padrao, arquivo_padrao)

    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        caminho = filedialog.askopenfilename(
            title="Selecione o importador principal",
            initialdir=diretorio_padrao,
            initialfile=arquivo_padrao,
            filetypes=[("Python", "*.py"), ("Todos os arquivos", "*.*")],
        )
        root.destroy()
    except Exception as e:
        _log(f"Popup indisponivel ({e}).")
        caminho = ""

    caminho = (caminho or "").strip()
    if caminho:
        return caminho

    if os.path.exists(caminho_padrao):
        _log(f"Nenhum arquivo selecionado. Usando padrao: {caminho_padrao}")
        return caminho_padrao

    return ""


def caminho_checkpoint(caminho_script: str) -> str:
    return os.path.join(os.path.dirname(caminho_script), CHECKPOINT_NOME_ARQUIVO)


def ler_checkpoint(caminho_script: str):
    caminho = caminho_checkpoint(caminho_script)
    try:
        if not os.path.exists(caminho):
            return None
        with open(caminho, "r", encoding="utf-8") as f:
            linha = (f.readline() or "").strip()
        partes = linha.split("|")
        if len(partes) < 2:
            return None
        codigo = partes[0].strip()
        estab = partes[1].strip()
        if not codigo:
            return None
        return (codigo, estab or "1")
    except Exception as e:
        _log(f"Falha ao ler checkpoint: {e}")
        return None


def limpar_checkpoint(caminho_script: str) -> None:
    caminho = caminho_checkpoint(caminho_script)
    try:
        if os.path.exists(caminho):
            os.remove(caminho)
    except Exception as e:
        _log(f"Falha ao limpar checkpoint: {e}")


def _nome(ctrl) -> str:
    try:
        return (getattr(ctrl, "Name", "") or "").strip()
    except Exception:
        return ""


def _tipo(ctrl) -> str:
    try:
        return getattr(ctrl, "ControlTypeName", "") or ""
    except Exception:
        return ""


def _parent(ctrl):
    try:
        return ctrl.GetParentControl()
    except Exception:
        return None


def _iter_desc(ctrl):
    fila = deque([ctrl])
    while fila:
        atual = fila.popleft()
        yield atual
        try:
            for ch in atual.GetChildren():
                fila.append(ch)
        except Exception:
            pass


def _iter_buttons(ctrl):
    for node in _iter_desc(ctrl):
        if _tipo(node) == "ButtonControl":
            yield node


def _dialog_ancestor(node):
    cur = node
    while cur:
        if _tipo(cur) in ("DialogControl", "WindowControl"):
            return cur
        cur = _parent(cur)
    return node


def _iter_top_windows():
    root = uia.GetRootControl()
    try:
        for w in root.GetChildren():
            if _tipo(w) == "WindowControl":
                yield w
    except Exception:
        return


def _find_window_by_title_substr(sub: str):
    alvo = _normalizar_txt(sub)
    for w in _iter_top_windows():
        if alvo in _normalizar_txt(_nome(w)):
            return w
    return None


def _window_rect(ctrl):
    try:
        r = ctrl.BoundingRectangle
        left, top = int(r.left), int(r.top)
        right, bottom = int(r.right), int(r.bottom)
        return left, top, right, bottom
    except Exception:
        return None


def _collect_texts(ctrl, max_depth=8):
    out = []
    fila = deque([(ctrl, 0)])
    while fila:
        node, depth = fila.popleft()
        if depth > max_depth:
            continue
        try:
            nm = _nome(node)
            if nm:
                out.append(nm)
            for ch in node.GetChildren():
                fila.append((ch, depth + 1))
        except Exception:
            pass
    return out


def _invoke_or_default(ctrl) -> bool:
    try:
        inv = ctrl.GetInvokePattern()
        if inv:
            inv.Invoke()
            return True
    except Exception:
        pass
    try:
        lg = ctrl.GetLegacyIAccessiblePattern()
        if lg:
            lg.DoDefaultAction()
            return True
    except Exception:
        pass
    return False


def _janela_parece_crash_fiscal(win) -> bool:
    if FISCAL_TITLE not in _normalizar_txt(_nome(win)):
        return False

    txt_norm = _normalizar_txt("\n".join(_collect_texts(win, max_depth=8)))
    if ("access violation" in txt_norm) or ("sef.exe" in txt_norm) or ("read of address" in txt_norm):
        return True

    rect = _window_rect(win)
    if not rect:
        return False
    left, top, right, bottom = rect
    largura = max(1, right - left)
    altura = max(1, bottom - top)

    # Dialogo de erro costuma ser pequeno e com botao OK unico.
    if largura <= 1200 and altura <= 700:
        for btn in _iter_buttons(win):
            if _normalizar_txt(_nome(btn)) == "ok":
                return True
    return False


def _detectar_dialogo_crash_fiscal():
    for w in _iter_top_windows():
        try:
            if _janela_parece_crash_fiscal(w):
                return w
        except Exception:
            continue
    return None


def _fechar_janela(win) -> bool:
    try:
        wp = win.GetWindowPattern()
        if wp:
            wp.Close()
            return True
    except Exception:
        pass
    for btn in _iter_buttons(win):
        nm = _normalizar_txt(_nome(btn))
        if nm in ("fechar", "close", "ok"):
            if _invoke_or_default(btn):
                return True
    try:
        win.SetFocus()
        time.sleep(0.05)
        pag.hotkey("alt", "f4")
        return True
    except Exception:
        pass
    return False


def _fechar_todas_janelas_fiscal(timeout_total=5.0):
    fim = time.time() + timeout_total
    while time.time() < fim:
        fiscais = [w for w in _iter_top_windows() if FISCAL_TITLE in _normalizar_txt(_nome(w))]
        if not fiscais:
            return True
        fechou_alguma = False
        for win in fiscais:
            if _fechar_janela(win):
                fechou_alguma = True
        if not fechou_alguma:
            break
        time.sleep(0.4)
    fiscais_restantes = [w for w in _iter_top_windows() if FISCAL_TITLE in _normalizar_txt(_nome(w))]
    return len(fiscais_restantes) == 0


def _botoes_confirmacao_em_janela(win):
    pares = []
    for btn in _iter_buttons(win):
        nm = _normalizar_txt(_nome(btn))
        if nm:
            pares.append((nm, btn))
    return pares


def _janela_tem_contexto_de_erro_ou_aviso(win) -> bool:
    nm = _normalizar_txt(_nome(win))
    if any(chave in nm for chave in (NOME_ALERTA, "erro", "error", "mensagem", "aviso")):
        return True

    if FISCAL_TITLE in nm:
        rect = _window_rect(win)
        if rect:
            left, top, right, bottom = rect
            largura = max(1, right - left)
            altura = max(1, bottom - top)
            # Dialogos do Fiscal costumam ser menores que a janela principal.
            if largura <= 1300 and altura <= 800:
                return True

    txt = _normalizar_txt("\n".join(_collect_texts(win, max_depth=8)))
    return any(chave in txt for chave in ("access violation", "read of address", "sef.exe", "erro", "error", "aviso"))


def _confirmar_janela_por_botoes(win) -> bool:
    prioridades = ("sim", "ok", "yes", "confirmar", "fechar", "close", "encerrar")
    btns = _botoes_confirmacao_em_janela(win)
    if not btns:
        return False

    for alvo in prioridades:
        for nome_btn, btn in btns:
            if alvo in nome_btn:
                if _invoke_or_default(btn):
                    return True
    return False


def confirmar_todos_erros_e_avisos_antes_de_fechar(timeout_total=12.0, interval=0.25) -> int:
    """
    Confirma em lote todos os dialogs/avisos (Sim/OK/Fechar) antes de fechar o Fiscal.
    """
    confirmados = 0
    fim = time.time() + timeout_total
    ultimo_confirmado = 0.0

    while time.time() < fim:
        atuou = False
        for win in list(_iter_top_windows()):
            try:
                if not _janela_tem_contexto_de_erro_ou_aviso(win):
                    continue
                if _confirmar_janela_por_botoes(win):
                    confirmados += 1
                    ultimo_confirmado = time.time()
                    atuou = True
                    _log(f"Dialogo/aviso confirmado: {_nome(win) or '(sem titulo)'}")
                    time.sleep(0.2)
            except Exception:
                continue

        if atuou:
            continue

        if confirmados > 0 and (time.time() - ultimo_confirmado) > 1.0:
            break
        time.sleep(interval)

    return confirmados


def _acionar_botao_fiscal_no_gerenciador() -> bool:
    ger = _find_window_by_title_substr(GERENCIADOR_TITLE)
    if not ger:
        _log("Janela 'Gerenciador de Sistemas' nao encontrada.")
        return False

    try:
        ger.SetFocus()
    except Exception:
        pass
    time.sleep(0.2)

    rect = _window_rect(ger)
    if rect:
        left, top, right, bottom = rect
        width = max(1, right - left)
        height = max(1, bottom - top)
        x = left + int(round(width * BOTAO_FISCAL_RATIO_X))
        y = top + int(round(height * BOTAO_FISCAL_RATIO_Y))
    else:
        x, y = BOTAO_FISCAL_ABS_X, BOTAO_FISCAL_ABS_Y

    # Primeiro clique simples; fallback duplo.
    pag.moveTo(x, y, duration=0.12)
    pag.click(x, y)
    for _ in range(8):
        time.sleep(1)
        fiscais = [w for w in _iter_top_windows() if FISCAL_TITLE in _normalizar_txt(_nome(w))]
        if fiscais:
            _log(f"Botao 'Fiscal' acionado no Gerenciador em {x},{y}.")
            return True

    pag.doubleClick(x, y)
    for _ in range(6):
        time.sleep(1)
        fiscais = [w for w in _iter_top_windows() if FISCAL_TITLE in _normalizar_txt(_nome(w))]
        if fiscais:
            _log(f"Botao 'Fiscal' acionado com duplo clique em {x},{y}.")
            return True

    _log(f"Nao foi possivel abrir o Fiscal pelo Gerenciador (coord {x},{y}).")
    return False


def reabrir_fiscal_via_gerenciador(motivo: str) -> bool:
    """
    Regra geral de intervencao do watchdog:
      - fecha o Fiscal
      - reabre pelo Gerenciador de Sistemas
    """
    _log(f"Intervencao watchdog ({motivo}): confirmando erros/avisos, fechando Fiscal e reabrindo pelo Gerenciador...")

    confirmados = confirmar_todos_erros_e_avisos_antes_de_fechar(timeout_total=12.0, interval=0.25)
    if confirmados > 0:
        _log(f"Dialogs confirmados antes do fechamento do Fiscal: {confirmados}")

    fechado = False
    for rodada in range(1, 4):
        _fechar_todas_janelas_fiscal(timeout_total=7.0)
        confirmar_todos_erros_e_avisos_antes_de_fechar(timeout_total=4.0, interval=0.25)
        fiscais_restantes = [w for w in _iter_top_windows() if FISCAL_TITLE in _normalizar_txt(_nome(w))]
        if not fiscais_restantes:
            fechado = True
            break
        _log(f"Ainda existe janela do Fiscal apos tentativa {rodada}/3 de fechamento.")
        time.sleep(0.5)

    if not fechado:
        _log("Nao foi possivel encerrar completamente o Fiscal; abortando reabertura para evitar selecao de empresa.")
        return False

    ok = _acionar_botao_fiscal_no_gerenciador()
    if ok:
        _log("Fiscal reaberto com sucesso pelo Gerenciador.")
        return True
    _log("Falha ao reabrir Fiscal pelo Gerenciador.")
    return False


def tratar_crash_access_violation_fiscal() -> bool:
    """
    Trata o erro:
      Access violation at address ... in module 'sef.exe' ...
    Fluxo:
      1) fecha dialogo/janelas Fiscal
      2) ativa Gerenciador de Sistemas
      3) aciona botao cinza 'Fiscal'
    """
    dlg = _detectar_dialogo_crash_fiscal()
    if not dlg:
        return False

    _log("Detectado crash do Fiscal (Access violation). Iniciando recuperacao...")

    # Fecha o dialogo de erro (normalmente botao OK).
    _fechar_janela(dlg)
    time.sleep(0.5)
    ok_reabrir = reabrir_fiscal_via_gerenciador("crash access violation")
    if ok_reabrir:
        _log("Recuperacao do Fiscal concluida.")
        return True

    _log("Recuperacao do Fiscal falhou ao acionar o Gerenciador.")
    return False


def localizar_aviso_do_sistema(max_depth: int = 10):
    root = uia.GetRootControl()
    fila = deque([(root, 0)])

    while fila:
        node, depth = fila.popleft()
        if depth > max_depth:
            continue

        try:
            tipo = _tipo(node)
            nome = _normalizar_txt(_nome(node))
            if tipo in ("DialogControl", "WindowControl") and nome == NOME_ALERTA:
                return node
            if tipo == "TitleBarControl" and nome == NOME_ALERTA:
                return _dialog_ancestor(node)
        except Exception:
            pass

        try:
            for ch in node.GetChildren():
                fila.append((ch, depth + 1))
        except Exception:
            pass

    return None


def confirmar_aviso(ctrl_aviso):
    prioridades = ("sim", "ok", "yes", "confirmar")

    for prioridade in prioridades:
        try:
            for btn in _iter_buttons(ctrl_aviso):
                nome_btn = _normalizar_txt(_nome(btn))
                if prioridade in nome_btn:
                    try:
                        invoke = btn.GetInvokePattern()
                        if invoke:
                            invoke.Invoke()
                            return _nome(btn) or prioridade
                    except Exception:
                        try:
                            legacy = btn.GetLegacyIAccessiblePattern()
                            if legacy:
                                legacy.DoDefaultAction()
                                return _nome(btn) or prioridade
                        except Exception:
                            pass
        except Exception:
            pass

    return None


def confirmar_avisos_pendentes(timeout_total: float = 10.0, interval: float = 0.25) -> int:
    confirmados = 0
    fim = time.time() + timeout_total
    ultimo_confirmado = 0.0

    while time.time() < fim:
        aviso = localizar_aviso_do_sistema(max_depth=10)
        if not aviso:
            if confirmados > 0 and (time.time() - ultimo_confirmado) > 1.0:
                break
            time.sleep(interval)
            continue

        botao = confirmar_aviso(aviso)
        if botao:
            confirmados += 1
            ultimo_confirmado = time.time()
            _log(f"Aviso confirmado com: {botao}")
        else:
            _log("Aviso detectado, mas sem botao acionavel.")
            break
        time.sleep(0.3)

    return confirmados


def iniciar_importador(caminho_script: str, empresa):
    env = os.environ.copy()
    if empresa:
        env["IMPORTADOR_START_CODIGO"] = str(empresa[0])
        env["IMPORTADOR_START_ESTAB"] = str(empresa[1])
        _log(f"Reinicio direcionado para Empresa {empresa[0]} / Estab {empresa[1]}")
    else:
        env.pop("IMPORTADOR_START_CODIGO", None)
        env.pop("IMPORTADOR_START_ESTAB", None)
        _log("Inicio sem ponto de retomada (do comeco).")

    comando = [sys.executable, "-u", caminho_script]
    return subprocess.Popen(
        comando,
        cwd=os.path.dirname(caminho_script),
        env=env,
    )


def aguardar_termino(proc: subprocess.Popen) -> int:
    while True:
        if tratar_crash_access_violation_fiscal():
            try:
                if proc.poll() is None:
                    proc.terminate()
                    time.sleep(1.0)
                    if proc.poll() is None:
                        proc.kill()
            except Exception:
                pass
            return CODIGO_RECUPERADO_CRASH_FISCAL

        codigo = proc.poll()
        if codigo is not None:
            return codigo
        time.sleep(0.6)


def main():
    caminho_script = escolher_script_importador()
    if not caminho_script:
        _log("Nenhum script selecionado. Encerrando.")
        return

    if not os.path.exists(caminho_script):
        _log(f"Script nao encontrado: {caminho_script}")
        return

    _log(f"Script monitorado: {caminho_script}")
    empresa_retomada = None
    proc = None

    try:
        for tentativa in range(0, MAX_RESTARTS + 1):
            proc = iniciar_importador(caminho_script, empresa_retomada)
            codigo_saida = aguardar_termino(proc)

            if codigo_saida == CODIGO_RECUPERADO_CRASH_FISCAL:
                empresa_checkpoint = ler_checkpoint(caminho_script)
                if empresa_checkpoint:
                    empresa_retomada = empresa_checkpoint
                    _log(f"Retomada apos crash em Empresa {empresa_retomada[0]} / Estab {empresa_retomada[1]}")
                else:
                    _log("Checkpoint nao encontrado apos crash; reinicio sera do comeco.")
                    empresa_retomada = None

                if tentativa >= MAX_RESTARTS:
                    _log("Limite de reinicios atingido. Encerrando watchdog.")
                    return

                _log("Reiniciando importador apos recuperacao de crash do Fiscal.")
                time.sleep(RESTART_DELAY)
                continue

            if codigo_saida == 0:
                limpar_checkpoint(caminho_script)
                _log("Importador finalizado sem erro. Watchdog encerrado.")
                return

            _log(f"Importador encerrou com erro (saida={codigo_saida}). Verificando aviso do sistema...")
            avisos_confirmados = confirmar_avisos_pendentes(timeout_total=10.0, interval=0.25)

            if avisos_confirmados > 0:
                _log(f"Avisos confirmados apos falha: {avisos_confirmados}")
            else:
                _log("Nenhum aviso confirmavel apos a falha, mas watchdog fara reabertura padrao do Fiscal.")

            if not reabrir_fiscal_via_gerenciador("falha do importador"):
                _log("Nao foi possivel concluir a reabertura padrao do Fiscal. Encerrando watchdog.")
                return

            empresa_checkpoint = ler_checkpoint(caminho_script)
            if empresa_checkpoint:
                empresa_retomada = empresa_checkpoint
                _log(f"Retomada preparada em Empresa {empresa_retomada[0]} / Estab {empresa_retomada[1]}")
            else:
                _log("Checkpoint nao encontrado; reinicio sera do comeco.")
                empresa_retomada = None

            if tentativa >= MAX_RESTARTS:
                _log("Limite de reinicios atingido. Encerrando watchdog.")
                return

            _log(f"Reiniciando importador apos confirmar aviso ({avisos_confirmados}).")
            time.sleep(RESTART_DELAY)
    except KeyboardInterrupt:
        _log("Interrompido pelo usuario (Ctrl+C). Encerrando watchdog.")
        try:
            if proc and proc.poll() is None:
                proc.terminate()
                time.sleep(1.0)
                if proc.poll() is None:
                    proc.kill()
        except Exception:
            pass


if __name__ == "__main__":
    main()
