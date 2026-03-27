# -*- coding: utf-8 -*-
"""
Escrituração Contábil Digital (ECD)
"""

from __future__ import annotations

import datetime as _dt
import json
import os
import sys
import time
import tkinter as tk
import unicodedata

try:
    import pyautogui as pag
except Exception as exc:
    raise SystemExit("pyautogui é obrigatório para a automação de UI") from exc

try:
    from pywinauto import Desktop
except Exception as exc:
    raise SystemExit("pywinauto é obrigatório para detectar o aviso") from exc


# =====================
# Configuracoes
# =====================
NOME_SCRIPT = "4 - Escrituração Contábil Digital (ECD)"
BASE_DIR = r"W:\SPEDs\ECD\2025"
ANO_PASSADO = _dt.date.today().year - 1
ANO_PASSADO_CURTO = str(ANO_PASSADO)[-2:]
LOG_PATH = os.path.join(os.path.dirname(__file__), "log_ecd_4.txt")
STATUS_PATH = r"W:\\DOCUMENTOS ESCRITORIO\\INSTALACAO SISTEMA\\central-utils\\data\\ecd-status\\ecd_status.json"
AUTOMATIZADO_ERROS_BASE = r"W:\\DOCUMENTOS ESCRITORIO\\RH\\AUTOMATIZADO\\ECD\\erros_validacao"

# Coordenadas
CLIQUE_INICIAL = (401, 463)

BTN_ABRIR_MOVIMENTO_1 = (843, 449)
BTN_ABRIR_MOVIMENTO_2 = (813, 468)
BTN_ABRIR_MOVIMENTO_3 = (1095, 596)
BTN_ABRIR_MOVIMENTO_4 = (1028, 595)
BTN_ABRIR_MOVIMENTO_5 = (1217, 400)
BTN_ABRIR_MOVIMENTO_6 = (1257, 363)
BTN_ABRIR_MOVIMENTO_7 = (724, 476)
BTN_ABRIR_MOVIMENTO_8 = (826, 500)
BTN_ABRIR_MOVIMENTO_9 = (694, 492)

BTN_POS_1 = (541, 571)
BTN_POS_2 = (173, 28)
BTN_POS_3 = (232, 355)
BTN_POS_4 = (771, 707)
BTN_POS_5 = (796, 749)
BTN_POS_6 = (1224, 404)
BTN_POS_7 = (1265, 369)

CAMPO_ANO_CALENDARIO = (766, 312)
CAMPO_DIRETORIO_DESTINO = (948, 358)

BTN_GERAR_DMPL = (1060, 533)
BTN_GERAR_SOMENTE_CONTAS_SALDO_MOV = (739, 555)
BTN_GERAR_REGISTROS_PLANO_REFERENCIAL = (737, 534)

BTN_ADICIONAR_ARQUIVOS_RTF = (1059, 555)
BTN_ADICIONAR_ITEM_RTF = (1123, 466)
CAMPO_TIPO_DOCUMENTO = (792, 434)
BTN_SELECIONAR_ARQUIVO_RTF = (880, 472)
BTN_CONTINUAR_RTF = (1229, 424)

ABA_DEMONSTRACOES_CONTABEIS = (840, 287)
BTN_GERAR_DEMONSTRACOES_CONTABEIS = (589, 313)
BTN_VALIDAR = (1283, 347)
BTN_VALIDAR_V2 = (1302, 359)
BTN_VOLTAR_VALIDACAO = (1301, 388)
BTN_OK_VALIDACAO = (1284, 293)

# Tempos
WAIT_START = 2
WAIT_CURTO = 0.5
WAIT_POS_DIRETORIO = 0.5
TIMEOUT_AVISO = 90
INTERVALO_VERIFICACAO_VALIDACAO = 10
INTERVALO_TENTAR_OK_APOS_VALIDAR_V2 = 30
INTERVALO_RETENTATIVA_OK_FINAL = 120

AVISO_TITULO = "Aviso do Sistema"
AVISO_CLASS = "#32770"
JANELA_VALIDACAO_TEXTO = "Validacao de Regras da ECD"

NOME_RTF_NOTA_EXPLICATIVA = "Nota_explicativa.rtf"
NOME_RTF_DFC = "DFC.rtf"

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

def _abrir_janela_movimento() -> None:
    _log("Bind abrir movimento: segurar ALT, apertar M, apertar F, soltar ALT")
    pag.keyDown("alt")
    pag.press("m")
    _sleep(0.1, "intervalo entre M e F com ALT pressionado")
    pag.press("f")
    pag.keyUp("alt")
    _sleep(0.5, "delay apos bind abrir movimento")
    pag.write(ANO_PASSADO_CURTO, interval=0.02)
    pag.press("enter")
    _sleep(0.5, "delay apos digitar ano curto e confirmar")

def _abrir_movimento() -> None:
    _log("Mini fluxo: abrir movimento")
    _abrir_janela_movimento()

    cliques = [
        BTN_ABRIR_MOVIMENTO_1,
        BTN_ABRIR_MOVIMENTO_2,
        BTN_ABRIR_MOVIMENTO_3,
        BTN_ABRIR_MOVIMENTO_4,
        BTN_ABRIR_MOVIMENTO_5,
        BTN_ABRIR_MOVIMENTO_6,
        BTN_ABRIR_MOVIMENTO_7,
        BTN_ABRIR_MOVIMENTO_8,
        BTN_ABRIR_MOVIMENTO_9,
    ]
    for idx, pos in enumerate(cliques, start=1):
        _click(pos, f"abrir movimento {idx}")
        _sleep(0.5, "delay entre botoes do abrir movimento")

def _countdown(segundos: int) -> None:
    for restante in range(segundos, 0, -1):
        _log(f"Iniciando em {restante}s...")
        time.sleep(1)


def _set_clipboard(texto: str) -> None:
    root = tk.Tk()
    root.withdraw()
    root.clipboard_clear()
    root.clipboard_append(texto)
    root.update()
    root.destroy()


def _digitar_caminho(caminho: str) -> None:
    pag.hotkey("ctrl", "a")
    pag.press("backspace")
    _set_clipboard(caminho)
    pag.hotkey("ctrl", "v")


def _carregar_status() -> dict:
    try:
        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        _log(f"Arquivo de status nao encontrado: {STATUS_PATH}")
    except json.JSONDecodeError:
        _log("Arquivo de status com JSON invalido.")
    except Exception as exc:
        _log(f"Falha ao ler status: {exc}")
    return {}

def _salvar_status(status: dict) -> None:
    try:
        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(status, f, ensure_ascii=False, indent=2)
    except Exception as exc:
        _log(f"Falha ao salvar status: {exc}")

def _registrar_erro_status(empresa: str, mensagem: str, grave: bool = False) -> None:
    status = _carregar_status()
    companies = status.get("companies") or {}
    codigo_encontrado = None
    dados_encontrados = None
    def _norm(txt: str) -> str:
        txt = unicodedata.normalize("NFD", txt or "")
        txt = "".join(ch for ch in txt if unicodedata.category(ch) != "Mn")
        txt = "".join(ch.lower() for ch in txt if ch.isalnum() or ch.isspace())
        return " ".join(txt.split())
    if isinstance(companies, dict):
        alvo = _norm(empresa)
        for codigo, dados in companies.items():
            if not isinstance(dados, dict):
                continue
            nome_json = str(dados.get("name") or "").strip()
            nome_norm = _norm(nome_json)
            if nome_norm == alvo or nome_norm in alvo or alvo in nome_norm:
                codigo_encontrado = codigo
                dados_encontrados = dados
                break
    if not codigo_encontrado or dados_encontrados is None:
        _log(f"ERRO nao registrado: empresa nao encontrada no status ({empresa})")
        return

    dados_encontrados["erro"] = "Y"
    if grave:
        dados_encontrados["erroGrave"] = "Y"
    dados_encontrados["erroMsg"] = mensagem
    dados_encontrados["erroAt"] = _dt.datetime.now().isoformat()
    companies[codigo_encontrado] = dados_encontrados
    status["companies"] = companies
    _salvar_status(status)
    _log(f"ERRO registrado no status: {codigo_encontrado} - {mensagem}")


def _obter_flags_empresa(nome: str, argv: list[str]) -> tuple[bool, str]:
    dfc_arg = None
    simples_arg = None
    for a in argv[1:]:
        low = a.strip().lower()
        if low == "--dfc":
            dfc_arg = True
        elif low == "--no-dfc":
            dfc_arg = False
        elif low.startswith("--simples="):
            simples_arg = a.split("=", 1)[1].strip()
    if dfc_arg is not None or simples_arg:
        return bool(dfc_arg), (simples_arg or "Normal")
    status = _carregar_status()
    companies = status.get("companies") or {}
    if isinstance(companies, dict):
        for _codigo, dados in companies.items():
            if not isinstance(dados, dict):
                continue
            nome_json = str(dados.get("name") or "").strip()
            if nome_json.lower() == nome.strip().lower():
                dfc = bool(dados.get("dfc", False))
                simples = str(dados.get("simples") or "Normal").strip()
                return dfc, simples
    _log(f"Empresa nao localizada no status: {nome}")
    resp_dfc = input("Empresa tem DFC? (s/n): ").strip().lower()
    dfc = resp_dfc in {"s", "sim", "y", "yes"}
    resp_simples = input("Empresa é Simples? (s/n): ").strip().lower()
    simples = "Simples" if resp_simples in {"s", "sim", "y", "yes"} else "Normal"
    return dfc, simples


def _modo_pre_validar(argv: list[str]) -> bool:
    return "--pre-validar" in [a.strip().lower() for a in argv]

def _modo_teste_prefluxo(argv: list[str]) -> bool:
    return "--teste-prefluxo" in [a.strip().lower() for a in argv]

def _modo_teste_posfluxo(argv: list[str]) -> bool:
    return "--teste-posfluxo" in [a.strip().lower() for a in argv]

def _modo_parar_antes_ok_final(argv: list[str]) -> bool:
    return "--parar-antes-ok-final" in [a.strip().lower() for a in argv]

def _modo_somente_movimento(argv: list[str]) -> bool:
    return "--somente-movimento" in [a.strip().lower() for a in argv]

def _normalizar_texto(txt: str) -> str:
    txt = unicodedata.normalize("NFD", txt or "")
    txt = "".join(ch for ch in txt if unicodedata.category(ch) != "Mn")
    txt = "".join(ch.lower() for ch in txt if ch.isalnum() or ch.isspace())
    return " ".join(txt.split())

def _esperar_aviso(timeout: int) -> bool:
    t0 = time.time()
    while time.time() - t0 < timeout:
        restante = max(0, int(timeout - (time.time() - t0)))
        if restante % 5 == 0:
            _log(f"Aguardando aviso... restante {restante}s")
        try:
            win = Desktop(backend="win32").window(class_name=AVISO_CLASS, title=AVISO_TITULO)
            if win.exists(timeout=0.2):
                _log("Aviso do Sistema localizado.")
                return True
        except Exception:
            pass
        time.sleep(5)
    _log(f"Timeout aguardando aviso ({timeout}s).")
    return False

def _aviso_presente() -> bool:
    try:
        win = Desktop(backend="win32").window(class_name=AVISO_CLASS, title=AVISO_TITULO)
        return win.exists(timeout=0.2)
    except Exception:
        return False

def _aviso_texto() -> str:
    try:
        win = Desktop(backend="win32").window(class_name=AVISO_CLASS, title=AVISO_TITULO)
        if not win.exists(timeout=0.2):
            return ""
        textos = []
        try:
            for ch in win.descendants():
                try:
                    if ch.friendly_class_name() == "Static":
                        t = ch.window_text()
                        if t:
                            textos.append(t.strip())
                except Exception:
                    continue
        except Exception:
            pass
        return " ".join(textos)
    except Exception:
        return ""

def _aviso_eh_periodo_encerrado() -> bool:
    texto = _aviso_texto().lower()
    if not texto:
        return False
    return (
        "nao foi possivel corrigir a classificacao" in texto
        or "não foi possível corrigir a classificação" in texto
        or "periodo encontra-se encerrado" in texto
        or "período encontra-se encerrado" in texto
    )

def _aviso_final_eh_sucesso(texto: str) -> bool:
    t = _normalizar_texto(texto)
    if not t:
        return False
    padroes_sucesso = (
        "concluido com sucesso",
        "concluida com sucesso",
        "conclusao com sucesso",
        "gerado com sucesso",
        "gerada com sucesso",
        "validado com sucesso",
        "validada com sucesso",
    )
    return any(p in t for p in padroes_sucesso)

def _esperar_aviso_infinito() -> None:
    t0 = time.time()
    while True:
        try:
            win = Desktop(backend="win32").window(class_name=AVISO_CLASS, title=AVISO_TITULO)
            if win.exists(timeout=0.2):
                _log("Aviso do Sistema localizado.")
                return
        except Exception:
            pass
        if int(time.time() - t0) % 60 == 0:
            _log("Aguardando aviso final...")
        time.sleep(5)

def _esperar_aviso_final_com_retentativa_ok() -> None:
    t0 = time.time()
    proxima_retentativa = t0 + INTERVALO_RETENTATIVA_OK_FINAL
    while True:
        if _aviso_presente():
            _log("Aviso do Sistema localizado.")
            return
        agora = time.time()
        if agora >= proxima_retentativa:
            _log("Aviso final ainda nao apareceu. Repetindo clique em OK final.")
            _click(BTN_OK_VALIDACAO, "ok final (retentativa)")
            proxima_retentativa = agora + INTERVALO_RETENTATIVA_OK_FINAL
        if int(agora - t0) % 60 == 0:
            _log("Aguardando aviso final...")
        time.sleep(5)

def _obter_janela_validacao():
    alvo = _normalizar_texto(JANELA_VALIDACAO_TEXTO)
    try:
        for win in Desktop(backend="win32").windows():
            try:
                titulo = _normalizar_texto(win.window_text())
                if "validacao" in titulo and "ecd" in titulo:
                    return win
                if titulo == alvo:
                    return win
            except Exception:
                continue
    except Exception:
        return None
    return None

def _coletar_assinatura_validacao() -> str:
    win = _obter_janela_validacao()
    if win is not None:
        try:
            rect = win.rectangle()
            largura = max(1, int(rect.right - rect.left))
            altura = max(1, int(rect.bottom - rect.top))
            # Captura somente a area da grade/lista de regras (evita ruido do restante da tela).
            x = int(rect.left + (largura * 0.02))
            y = int(rect.top + (altura * 0.36))
            w = max(1, int(largura * 0.78))
            h = max(1, int(altura * 0.40))
            img = pag.screenshot(region=(x, y, w, h))
            img = img.resize((64, 36)).convert("L")
            return img.tobytes().hex()
        except Exception:
            pass
    return ""

def _janela_validacao_tem_regra_visivel() -> bool:
    win = _obter_janela_validacao()
    if win is None:
        return False
    fixos = {
        "Validação de Regras da ECD",
        "Validacao de Regras da ECD",
        "É necessário corrigir todos os erros e inconsistências listados abaixo",
        "E necessario corrigir todos os erros e inconsistencias listados abaixo",
        "Regra",
        "FAQ",
        "Validar",
        "Sair",
        "Novo",
        "Menu",
    }
    fixos_norm = {_normalizar_texto(t) for t in fixos}
    try:
        for ch in win.descendants():
            try:
                t = (ch.window_text() or "").strip()
            except Exception:
                continue
            if not t:
                continue
            tn = _normalizar_texto(t)
            if not tn or tn in fixos_norm:
                continue
            # Sobrou texto que nao e parte fixa da tela: trata como regra/erro visivel.
            return True
    except Exception:
        return False
    return False

def _esperar_resultado_validacao(timeout: int, intervalo_verificacao: int) -> str:
    t0 = time.time()
    proxima_verificacao = t0 + intervalo_verificacao
    assinatura_inicial = _coletar_assinatura_validacao()
    if assinatura_inicial:
        _log("Assinatura inicial da janela de validacao capturada.")

    while time.time() - t0 < timeout:
        if _aviso_presente():
            _log("Aviso do Sistema localizado durante a validacao.")
            return "aviso"

        agora = time.time()
        if agora >= proxima_verificacao:
            assinatura_atual = _coletar_assinatura_validacao()
            minutos = int((agora - t0) // 60)
            if assinatura_inicial and assinatura_atual and assinatura_atual != assinatura_inicial:
                _log(f"Mudanca detectada na janela de validacao apos {minutos} minuto(s).")
                return "mudou"
            if not assinatura_inicial and assinatura_atual:
                assinatura_inicial = assinatura_atual
                _log("Assinatura da janela de validacao capturada durante a espera.")
            else:
                _log(f"Sem mudanca na janela de validacao apos {minutos} minuto(s).")
            proxima_verificacao = agora + intervalo_verificacao

        time.sleep(5)

    _log(f"Timeout aguardando resultado da validacao ({timeout}s).")
    return "timeout"

def _esperar_resultado_validacao_com_retentativa_ok(
    intervalo_verificacao: int,
    intervalo_tentar_ok: int,
) -> str:
    t0 = time.time()
    proxima_verificacao = t0 + intervalo_verificacao
    proxima_tentativa_ok = t0 + intervalo_tentar_ok
    assinatura_inicial = _coletar_assinatura_validacao()
    if assinatura_inicial:
        _log("Assinatura inicial da janela de validacao capturada.")

    while True:
        if _aviso_presente():
            _log("Aviso do Sistema localizado durante a validacao.")
            return "aviso"

        agora = time.time()
        if agora >= proxima_verificacao:
            assinatura_atual = _coletar_assinatura_validacao()
            minutos = int((agora - t0) // 60)
            if assinatura_inicial and assinatura_atual and assinatura_atual != assinatura_inicial:
                _log(f"Mudanca detectada na janela de validacao apos {minutos} minuto(s).")
                return "mudou"
            if not assinatura_inicial and assinatura_atual:
                assinatura_inicial = assinatura_atual
                _log("Assinatura da janela de validacao capturada durante a espera.")
            else:
                _log(f"Sem mudanca na janela de validacao apos {minutos} minuto(s).")
            proxima_verificacao = agora + intervalo_verificacao

        if agora >= proxima_tentativa_ok:
            _log("Sem aviso apos validar v2. Tentando clicar em OK final novamente.")
            _click(BTN_OK_VALIDACAO, "ok final (retentativa apos validar v2)")
            proxima_tentativa_ok = agora + intervalo_tentar_ok

        time.sleep(5)

def _esperar_resultado_validacao_ate_aparecer(
    timeout: int,
    intervalo_verificacao: int,
    assinatura_inicial: str = "",
) -> str:
    t0 = time.time()
    deadline = t0 + timeout
    proxima_verificacao = t0 + intervalo_verificacao
    if assinatura_inicial:
        _log("Assinatura inicial da janela de validacao capturada.")

    while True:
        if time.time() >= deadline:
            _log(f"Timeout aguardando resultado da validacao ({timeout}s).")
            return "timeout"
        if _aviso_presente():
            _log("Aviso do Sistema localizado durante a validacao.")
            return "aviso"
        if _janela_validacao_tem_regra_visivel():
            _log("Regra/erro visivel detectado na janela de validacao.")
            return "mudou"

        agora = time.time()
        if agora >= proxima_verificacao:
            assinatura_atual = _coletar_assinatura_validacao()
            minutos = int((agora - t0) // 60)
            if assinatura_inicial and assinatura_atual and assinatura_atual != assinatura_inicial:
                _log(f"Mudanca detectada na janela de validacao apos {minutos} minuto(s).")
                return "mudou"
            if not assinatura_inicial and assinatura_atual:
                assinatura_inicial = assinatura_atual
                _log("Assinatura da janela de validacao capturada durante a espera.")
            else:
                _log(f"Sem mudanca na janela de validacao apos {minutos} minuto(s).")
            proxima_verificacao = agora + intervalo_verificacao
        time.sleep(5)

def _destino_screenshot_erro(empresa: str) -> str:
    agora = _dt.datetime.now()
    ano = agora.strftime("%Y")
    mes = agora.strftime("%m")
    return os.path.join(AUTOMATIZADO_ERROS_BASE, ano, mes, empresa)

def _salvar_screenshot_erro(empresa: str) -> str:
    destino = _destino_screenshot_erro(empresa)
    os.makedirs(destino, exist_ok=True)
    caminho_img = os.path.join(destino, "erros registrados.png")
    img = pag.screenshot()
    img.save(caminho_img, format="PNG")
    _log(f"Screenshot salva: {caminho_img}")
    return caminho_img

def _salvar_screenshot_erro_grave(empresa: str) -> str:
    destino = _destino_screenshot_erro(empresa)
    os.makedirs(destino, exist_ok=True)
    caminho_img = os.path.join(destino, "ERRO_GRAVE.png")
    img = pag.screenshot()
    img.save(caminho_img, format="PNG")
    _log(f"Screenshot salva: {caminho_img}")
    return caminho_img

def _fechar_movimento() -> None:
    _log("Fechar movimento: abrir janela via ALT+M+F")
    _abrir_janela_movimento()
    _click(BTN_POS_1, "fechar movimento 1")
    _sleep(0.5, "delay entre botoes")
    _click(BTN_POS_2, "fechar movimento 2")
    _sleep(0.5, "delay entre botoes")
    _click(BTN_POS_3, "fechar movimento 3")
    _sleep(0.5, "delay entre botoes")
    pag.write(str(ANO_PASSADO), interval=0.02)
    _sleep(0.5, "delay entre botoes")
    _click(BTN_POS_4, "fechar movimento 4")
    _sleep(0.5, "delay entre botoes")
    _click(BTN_POS_5, "fechar movimento 5")
    _sleep(0.5, "delay entre botoes")
    _click(BTN_POS_6, "fechar movimento 6")
    _sleep(0.5, "delay entre botoes")
    _click(BTN_POS_7, "fechar movimento 7")

# =====================
# Fluxo principal
# =====================

def main() -> int:
    empresa = _obter_empresa(sys.argv)
    if not empresa:
        _log("Empresa vazia. Encerrando.")
        return 1

    teste_prefluxo = _modo_teste_prefluxo(sys.argv)
    teste_posfluxo = _modo_teste_posfluxo(sys.argv)
    somente_movimento = _modo_somente_movimento(sys.argv)
    if teste_prefluxo:
        _log("Modo teste-prefluxo ativo. Ignorando flags de DFC/Simples.")
        _countdown(3)
        _abrir_movimento()
        _log("Teste pre-fluxo finalizado.")
        return 0
    if teste_posfluxo:
        _log("Modo teste-posfluxo ativo.")
        _countdown(3)
        _fechar_movimento()
        _log("Teste pos-fluxo finalizado.")
        return 0
    if somente_movimento:
        _log("Modo somente-movimento ativo: abrindo e fechando movimento sem gerar ECD.")
        _abrir_movimento()
        _fechar_movimento()
        _log("Modo somente-movimento finalizado.")
        return 0

    dfc_ativo, simples = _obter_flags_empresa(empresa, sys.argv)
    pre_validar = _modo_pre_validar(sys.argv)
    parar_antes_ok_final = _modo_parar_antes_ok_final(sys.argv)
    tirou_foto_verificacao = False

    _abrir_movimento()

    _click(CLIQUE_INICIAL, "clique inicial")

    destino = _caminho_destino(empresa)
    os.makedirs(destino, exist_ok=True)
    caminho_rtf = os.path.join(destino, NOME_RTF_NOTA_EXPLICATIVA)

    _log(f"Iniciando: {NOME_SCRIPT}")
    _log(f"Empresa: {empresa}")
    _log(f"Diretorio destino: {destino}")
    _sleep(WAIT_START, "preparar a tela")
    _log(f"DFC: {'sim' if dfc_ativo else 'nao'} | Simples: {simples}")

    _log("ALT + MPG")
    _alt_sequencia("mpg")

    _sleep(WAIT_CURTO)
    _click(CAMPO_ANO_CALENDARIO, "ano calendario")
    pag.press("backspace", presses=4)
    pag.write(str(ANO_PASSADO), interval=0.02)
    pag.press("enter", presses=3)

    _click(CAMPO_DIRETORIO_DESTINO, "diretorio destino")
    _digitar_caminho(destino)
    _sleep(WAIT_POS_DIRETORIO, "pos-diretorio")

    _click(BTN_GERAR_DMPL, "gerar DMPL")
    _click(BTN_GERAR_SOMENTE_CONTAS_SALDO_MOV, "gerar contas com saldo ou movimentacao")
    if simples.strip().lower() != "normal":
        _click(BTN_GERAR_REGISTROS_PLANO_REFERENCIAL, "gerar registros plano referencial")
    else:
        _log("Empresa Normal: nao clicar em plano referencial.")

    _click(BTN_ADICIONAR_ARQUIVOS_RTF, "adicionar arquivos rtf")
    _sleep(WAIT_CURTO)
    _click(CAMPO_TIPO_DOCUMENTO, "tipo de documento")
    pag.write("10", interval=0.02)

    _click(BTN_SELECIONAR_ARQUIVO_RTF, "selecionar arquivo rtf")
    _sleep(WAIT_CURTO)
    _digitar_caminho(caminho_rtf)
    pag.press("enter")
    _click(BTN_ADICIONAR_ITEM_RTF, "adicionar rtf")

    if dfc_ativo:
        caminho_dfc = os.path.join(destino, NOME_RTF_DFC)
        _click(CAMPO_TIPO_DOCUMENTO, "tipo de documento (DFC)")
        pag.press("backspace", presses=2)
        pag.press("delete", presses=2)
        pag.write("2", interval=0.02)

        _click(BTN_SELECIONAR_ARQUIVO_RTF, "selecionar arquivo rtf (DFC)")
        _sleep(WAIT_CURTO)
        _digitar_caminho(caminho_dfc)
        pag.press("enter")
        _click(BTN_ADICIONAR_ITEM_RTF, "adicionar rtf (DFC)")
    _click(BTN_CONTINUAR_RTF, "continuar rtf")

    _click(ABA_DEMONSTRACOES_CONTABEIS, "aba demonstracoes contabeis")
    _click(BTN_GERAR_DEMONSTRACOES_CONTABEIS, "gerar demonstracoes contabeis")

    if pre_validar:
        _log("Modo pre-validar ativo. Encerrando antes da validacao.")
        return 0

    _click(BTN_VALIDAR, "validar")
    _sleep(0.5, "aguardar estado apos validar 1")
    assinatura_base_validacao = _coletar_assinatura_validacao()
    if assinatura_base_validacao:
        _log("Assinatura base da validacao capturada antes do validar v2.")
    _click(BTN_VALIDAR_V2, "validar v2")

    aviso_ok = False
    if _aviso_presente():
        _log("Aviso do Sistema ja presente. Fechando com ENTER.")
        aviso_ok = True
    else:
        _log(f"Aguardando aviso da validacao por ate {TIMEOUT_AVISO}s (sem OCR).")
        if _esperar_aviso(TIMEOUT_AVISO):
            aviso_ok = True
        else:
            _salvar_screenshot_erro(empresa)
            tirou_foto_verificacao = True
            _registrar_erro_status(empresa, "Resultado da validacao nao apareceu em ate 90 segundos.", grave=False)
            _log("Timeout da validacao (90s). Assumindo conclusao e seguindo fluxo.")
            _click(BTN_VOLTAR_VALIDACAO, "sair validacao")

    # Limpa referencia de OCR/assinatura da empresa ao fim do mini ciclo de validacao.
    assinatura_base_validacao = ""

    if aviso_ok:
        pag.press("enter")
        _sleep(WAIT_CURTO, "fechar aviso")
        _click(BTN_VOLTAR_VALIDACAO, "sair validacao")

    if parar_antes_ok_final:
        _log("Modo parar-antes-ok-final ativo. Encerrando antes do clique em OK final.")
        return 0

    _click(BTN_OK_VALIDACAO, "ok final")
    pag.press("enter", presses=2)
    _sleep(5, "aguardar 5s")
    pag.press("enter")
    _click(BTN_OK_VALIDACAO, "ok final 2")

    aviso_final_ok = False
    if _aviso_presente():
        _log("Segundo aviso localizado.")
        aviso_final_ok = True
    else:
        _esperar_aviso_final_com_retentativa_ok()
        aviso_final_ok = True
    texto_aviso_final = _aviso_texto()
    _log(f"Texto do aviso final: {texto_aviso_final or '(vazio)'}")

    if not _aviso_final_eh_sucesso(texto_aviso_final):
        if _aviso_eh_periodo_encerrado():
            _registrar_erro_status(empresa, "Periodo encerrado - nao foi possivel corrigir a classificacao.", grave=True)
            _salvar_screenshot_erro_grave(empresa)
        else:
            msg = f"Aviso final fora do padrao de sucesso: {texto_aviso_final or '(vazio)'}"
            _registrar_erro_status(empresa, msg, grave=False)
            _salvar_screenshot_erro(empresa)
        _log("Aviso final nao indica sucesso. Encerrando sem fechar o aviso para analise manual.")
        return 2

    pag.press("enter")
    _sleep(5, "aguardar 5s")
    _log("Finalizado.")

    if aviso_final_ok:
        pag.press("enter")
        _sleep(0.5, "pos-aviso final")

    _fechar_movimento()

    if tirou_foto_verificacao:
        _registrar_erro_status(empresa, "Screenshot de erro registrada na verificacao.", grave=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
