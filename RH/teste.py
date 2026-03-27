# -*- coding: utf-8 -*-
"""
Fluxo:
1) Tareffa (Playwright): login -> filtro -> exportar CSV -> salvar em ./downloads
2) CSV: montar lista única por (empresa, estabelecimento) a partir das colunas "Característica" e "codigoERP"
3) ERP (UIAutomation): ativar "Folha de Pagamento" -> trocar empresa (mesmo método do script antigo) -> gerar "Resumo da folha mensal"

Observações:
- Datas: sempre mês/ano anterior (MM/AAAA) e último dia do mês anterior (DDMMAAAA).
- Processamento: 2
- Classificação: 1
- Marcar: "Listar recibos de pró-labore", "Listar recibos de autônomos", "Imprimir dados em destaque"
"""

from __future__ import annotations

import csv
import datetime as dt
import re
import time
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

import pyautogui as pag
import pygetwindow as gw
import uiautomation as uia
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import ctypes
from pygetwindow import PyGetWindowException

# ===================== Datas (ERP) =====================
def ultimo_dia_mes_anterior(ref: dt.date | None = None) -> str:
    hoje = ref or dt.date.today()
    primeiro_do_mes = hoje.replace(day=1)
    ultimo = primeiro_do_mes - dt.timedelta(days=1)
    return f"{ultimo.day:02d}{ultimo.month:02d}{ultimo.year:04d}"  # DDMMAAAA


def obter_mes_ano_anterior_slash(ref: dt.date | None = None) -> str:
    hoje = ref or dt.date.today()
    primeiro_do_mes = hoje.replace(day=1)
    ultimo = primeiro_do_mes - dt.timedelta(days=1)
    return f"{ultimo.month:02d}/{ultimo.year:04d}"


# ===================== ERP (UIA) - Config =====================
FOLHA_TITLE_EXATO = "Folha de Pagamento"
WORKSPACE_NAME = "Espaço de trabalho"

MENU_RELATORIOS = "relatórios"
SUB_MENSAIS = "mensais"
ITEM_RESUMO_FOLHA = "resumo da folha mensal"

JANELA_RESUMO_FOLHA = "Resumo da folha de pagamento"

# >>> NOVOS IDs (descobertos no debug)
EDIT_MESANO_DE = "527206"       # campo esquerdo  (11/2025)
EDIT_MESANO_ATE = "592744"      # campo direito   (11/2025)

EDIT_PROCESSAMENTO = "658394"   # linha abaixo dos Mês/Ano
EDIT_CLASSIFICACAO = "396194"   # linha seguinte

# IDs antigos de checkbox/botão continuam
CHK_PROLABORE = "5244710"
CHK_AUTONOMOS = "1182160"
CHK_DESTAQUE = "3016776"
BTN_IMPRIMIR_ID = "1313120"

ROI_RIGHT_WIDTH = 560
ROI_TOP_OFFSET = 34
ROI_HEIGHT = 130
GRID_COLS = 10
GRID_ROWS = 6

# --- mapeamento por BoundingRectangle (coordenadas UIAutomation) ---

# Mês/Ano
RECT_MESANO_DE   = (770, 334, 816, 355)
RECT_MESANO_ATE  = (829, 334, 875, 355)

# Estabelecimentos (apenas mapeados, ainda não usados)
RECT_ESTAB_DE    = (770, 400, 816, 421)
RECT_ESTAB_ATE   = (849, 400, 895, 421)

# Checkboxes (apenas mapeados, para usar na versão final se quiser)
RECT_CHK_PROLABORE   = (926, 418, 1077, 435)
RECT_CHK_AUTONOMOS   = (926, 434, 1085, 451)
RECT_CHK_DESTAQUE    = (926, 466, 1082, 483)


# tamanhos mínimos/máximos dos campos de edição (para achar empresa/estab/data)
MIN_W, MAX_W = 18, 160
MIN_H, MAX_H = 14, 36

# ===================== UIA Helpers =====================
def _strip_accents(s: str) -> str:
    s = s or ""
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")


def _canon(s: str) -> str:
    return " ".join(_strip_accents(s).lower().split())


def normalize(s):
    try:
        return (s or "").strip().casefold()
    except Exception:
        return ""


def wait_until(fn, timeout=20.0, interval=0.25):
    end = time.time() + timeout
    while time.time() < end:
        try:
            v = fn()
            if v:
                return v
        except Exception:
            pass
        time.sleep(interval)
    return None


def bfs_find(root_ctrl, name_substr, types=(), max_depth=6):
    target = normalize(name_substr)
    q = [(root_ctrl, 0)]
    while q:
        node, depth = q.pop(0)
        if depth > max_depth:
            continue
        try:
            nm = normalize(getattr(node, "Name", ""))
            if target and target in nm:
                if not types or node.ControlTypeName in types:
                    return node
            for child in node.GetChildren():
                q.append((child, depth + 1))
        except Exception:
            pass
    return None


def bring_into_view(ctrl):
    try:
        sip = ctrl.GetScrollItemPattern()
        if sip:
            sip.ScrollIntoView()
    except Exception:
        pass
    try:
        ctrl.SetFocus()
    except Exception:
        pass


def uia_activate(ctrl, name_for_log="(controle)", prefer_invoke=True):
    try:
        bring_into_view(ctrl)
        if prefer_invoke:
            try:
                inv = ctrl.GetInvokePattern()
                if inv:
                    inv.Invoke()
                    return True
            except Exception:
                pass
        try:
            sel = ctrl.GetSelectionItemPattern()
            if sel:
                sel.Select()
                return True
        except Exception:
            pass
        try:
            ec = ctrl.GetExpandCollapsePattern()
            if ec:
                try:
                    ec.Expand()
                except Exception:
                    pass
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def rect_of(ctrl):
    try:
        r = ctrl.BoundingRectangle
        return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        return None

def center_of_rect(rect):
    l, t, r, b = rect
    return ((l + r) / 2.0, (t + b) / 2.0)


def find_control_near_rect(root_ctrl, target_rect, control_type_name, max_depth=10, max_dist=30):
    """
    Procura, a partir de root_ctrl, o controle do tipo `control_type_name`
    cujo BoundingRectangle é mais próximo de target_rect (pelo centro),
    com tolerância max_dist em distância Manhattan.
    """
    tx, ty = center_of_rect(target_rect)
    best = None
    best_dist = None

    q = [(root_ctrl, 0)]
    while q:
        node, d = q.pop(0)
        if d > max_depth:
            continue
        try:
            if node.ControlTypeName == control_type_name:
                r = rect_of(node)
                if r:
                    cx, cy = center_of_rect(r)
                    dist = abs(cx - tx) + abs(cy - ty)
                    if best is None or dist < best_dist:
                        best = node
                        best_dist = dist
            for ch in node.GetChildren():
                q.append((ch, d + 1))
        except Exception:
            pass

    if best is None:
        print(f"[LOG] [rect] Nenhum {control_type_name} candidato perto de {target_rect}.")
        return None

    print(f"[LOG] [rect] Melhor {control_type_name} para {target_rect}: "
          f"rect={rect_of(best)}, dist={best_dist:.1f}")
    if best_dist is not None and best_dist <= max_dist:
        return best

    print(f"[LOG] [rect] Distância {best_dist:.1f} acima da tolerância ({max_dist}). Ignorando.")
    return None


def find_edit_by_rect(main_win, rect_target, label_para_log):
    # tenta dentro da janela principal
    ctrl = find_control_near_rect(main_win, rect_target, "EditControl", max_depth=10, max_dist=40)
    if ctrl:
        print(f"[LOG] Campo {label_para_log} localizado via BoundingRectangle.")
        return ctrl
    # fallback global (se por algum motivo não achar via main_win)
    root = uia.GetRootControl()
    ctrl = find_control_near_rect(root, rect_target, "EditControl", max_depth=12, max_dist=40)
    if ctrl:
        print(f"[LOG] Campo {label_para_log} localizado GLOBALMENTE via BoundingRectangle.")
    return ctrl


def find_checkbox_by_rect(main_win, rect_target, label_para_log):
    ctrl = find_control_near_rect(main_win, rect_target, "CheckBoxControl", max_depth=10, max_dist=40)
    if ctrl:
        print(f"[LOG] Checkbox {label_para_log} localizado via BoundingRectangle.")
        return ctrl
    root = uia.GetRootControl()
    ctrl = find_control_near_rect(root, rect_target, "CheckBoxControl", max_depth=12, max_dist=40)
    if ctrl:
        print(f"[LOG] Checkbox {label_para_log} localizado GLOBALMENTE via BoundingRectangle.")
    return ctrl


def size_of(ctrl):
    r = rect_of(ctrl)
    if not r:
        return (0, 0)
    return (r[2] - r[0], r[3] - r[1])


def has_valuepattern(ctrl):
    try:
        return ctrl.GetValuePattern() is not None
    except Exception:
        return False


def is_edit_candidate(ctrl):
    try:
        if not ctrl.IsEnabled or not ctrl.IsKeyboardFocusable:
            return False
        if not has_valuepattern(ctrl):
            return False
        w, h = size_of(ctrl)
        return (MIN_W <= w <= MAX_W) and (MIN_H <= h <= MAX_H)
    except Exception:
        return False


def read_value(ctrl):
    try:
        return ctrl.GetValuePattern().Value or ""
    except Exception:
        return ""


def is_date_like(s):
    s = s or ""
    if "/" in s and re.search(r"\d{1,2}/\d{1,2}/\d{2,4}", s):
        return True
    if len(re.sub(r"\D", "", s)) >= 8:
        return True
    return False


def set_text(control: uia.Control, text: str):
    try:
        vp = control.GetValuePattern()
        vp.SetValue(text)
        return
    except Exception:
        pass
    control.SetFocus()
    uia.SendKeys("^a{DEL}")
    uia.SendKeys(text)

def ativar_e_maximizar():
    wins = [w for w in gw.getWindowsWithTitle(FOLHA_TITLE_EXATO) if w.visible]
    if not wins:
        print("❌ Janela do Fiscal não localizada.")
        return None
    w = max(wins, key=lambda x: x.width * x.height)
    if not w.isActive:
        w.activate()
    if not w.isMaximized:
        w.maximize()
    time.sleep(1.2)
    return w

# ===================== Relatório: Resumo da folha mensal =====================
def _find_workspace(main_win: uia.Control) -> Optional[uia.Control]:
    try:
        for ch in main_win.GetChildren():
            if ch.ControlTypeName == "PaneControl" and (getattr(ch, "Name", "") or "") == WORKSPACE_NAME:
                return ch
    except Exception:
        pass
    return None


def _find_menu_item_anywhere(root: uia.Control, text: str) -> Optional[uia.Control]:
    return bfs_find(
        root,
        text,
        types=("MenuItemControl", "TabItemControl", "ButtonControl", "ListItemControl", "TreeItemControl"),
        max_depth=10,
    )


def abrir_tela_resumo_folha():
    root = uia.GetRootControl()
    main_win = ativar_e_maximizar()
    if not main_win:
        raise RuntimeError(f"Janela '{FOLHA_TITLE_EXATO}' não localizada via UIA.")

    workspace = _find_workspace(main_win) or main_win

    print("[LOG] Procurando menu 'Relatórios'...")
    rel = wait_until(
        lambda: _find_menu_item_anywhere(workspace, MENU_RELATORIOS)
        or _find_menu_item_anywhere(root, MENU_RELATORIOS),
        timeout=20,
        interval=0.4,
    )
    if not rel:
        print("[LOG] NÃO encontrei menu 'Relatórios'.")
        raise RuntimeError("Menu 'Relatórios' não encontrado.")
    print("[LOG] Detectei menu 'Relatórios'.")
    uia_activate(rel, "menu 'Relatórios'")

    print("[LOG] Procurando submenu 'Mensais'...")
    mensais = wait_until(
        lambda: _find_menu_item_anywhere(workspace, SUB_MENSAIS)
        or _find_menu_item_anywhere(root, SUB_MENSAIS),
        timeout=20,
        interval=0.4,
    )
    if not mensais:
        print("[LOG] NÃO encontrei submenu 'Mensais'.")
        raise RuntimeError("Submenu 'Mensais' não encontrado.")
    print("[LOG] Detectei submenu 'Mensais'.")
    uia_activate(mensais, "submenu 'Mensais'")

    print("[LOG] Procurando item 'Resumo da folha mensal'...")
    item = wait_until(
        lambda: _find_menu_item_anywhere(workspace, ITEM_RESUMO_FOLHA)
        or _find_menu_item_anywhere(root, ITEM_RESUMO_FOLHA),
        timeout=20,
        interval=0.4,
    )
    if not item:
        print("[LOG] Não achei por ITEM_RESUMO_FOLHA, tentando nome alternativo 'Resumo da Folha de Pagamento'...")
        item = wait_until(
            lambda: _find_menu_item_anywhere(workspace, "Resumo da Folha de Pagamento")
            or _find_menu_item_anywhere(root, "Resumo da Folha de Pagamento"),
            timeout=8,
            interval=0.4,
        )
    if not item:
        print("[LOG] Tentando fallback em menus/árvore interativa para 'Resumo da Folha de Pagamento'...")
        item = wait_until(
            lambda: _find_menu_item_anywhere(workspace, "Resumo da Folha de Pagamento")
            or _find_menu_item_anywhere(root, "Resumo da Folha de Pagamento"),
            timeout=10,
            interval=0.4,
        )

    if not item:
        print("[LOG] NÃO encontrei item de menu 'Resumo da Folha'.")
        raise RuntimeError("Item do resumo não encontrado (nem menu, nem TreeItem).")

    print("[LOG] Detectei item 'Resumo da Folha'.")
    uia_activate(item, "item 'Resumo da Folha de Pagamento'")


def uia_activate(ctrl, name_for_log="(controle)", prefer_invoke=True):
    """Tenta Invoke/Select/Expand; agora loga o que está tentando acionar."""
    print(f"[LOG] Acionando {name_for_log}...")
    success = False
    try:
        bring_into_view(ctrl)

        if prefer_invoke:
            try:
                inv = ctrl.GetInvokePattern()
                if inv:
                    inv.Invoke()
                    success = True
            except Exception:
                pass

        if not success:
            try:
                sel = ctrl.GetSelectionItemPattern()
                if sel:
                    sel.Select()
                    success = True
            except Exception:
                pass

        if not success:
            try:
                ec = ctrl.GetExpandCollapsePattern()
                if ec:
                    try:
                        ec.Expand()
                    except Exception:
                        pass
                    success = True
            except Exception:
                pass
    except Exception:
        success = False

    if success:
        print(f"[LOG] {name_for_log} acionado com sucesso.")
    else:
        print(f"[WARN] Falha ao acionar {name_for_log}.")
    return success


def _find_resumo_container(main_win: uia.Control) -> Optional[uia.Control]:
    """
    Agora procura a tela do resumo GLOBALMENTE:
    1) tenta achar qualquer Edit com AutomationId EDIT_MESANO_DE
       em todo o desktop e sobe até um pai que tenha o botão Imprimir;
    2) se não achar, tenta pegar uma WindowControl cujo título contenha
       'resumo da folha';
    3) se nada disso der certo, retorna None (para o wait_until repetir).
    """
    root = uia.GetRootControl()

    # 1) procurar o campo Mês/Ano globalmente
    print("[LOG] [Resumo] Procurando campo Mês/Ano (EDIT_MESANO_DE) GLOBALMENTE...")
    ed_de = root.EditControl(AutomationId=EDIT_MESANO_DE)
    if ed_de.Exists(0, 0):
        print("[LOG] [Resumo] Achei EditControl EDIT_MESANO_DE; subindo pais até achar botão Imprimir...")
        node = ed_de
        for nivel in range(35):
            try:
                parent = node.GetParentControl()
                if not parent:
                    print(f"[LOG] [Resumo] Parei no nível {nivel}: parent None.")
                    break
                btn = parent.ButtonControl(AutomationId=BTN_IMPRIMIR_ID)
                if btn.Exists(0, 0):
                    print(f"[LOG] [Resumo] Container encontrado no nível {nivel} (tem botão Imprimir).")
                    return parent
                node = parent
            except Exception as e:
                print(f"[WARN] [Resumo] Erro subindo hierarquia: {e}")
                break

    print("[LOG] [Resumo] Não achei EDIT_MESANO_DE globalmente; tentando localizar janela por título...")

    # 2) procurar uma janela com nome contendo "resumo da folha"
    try:
        for w in root.GetChildren():
            if w.ControlTypeName != "WindowControl":
                continue
            name = (getattr(w, "Name", "") or "").strip().lower()
            if "resumo da folha" in name:
                print(f"[LOG] [Resumo] Janela candidata encontrada pelo título: {name!r}")
                return w
    except Exception as e:
        print(f"[WARN] [Resumo] Erro ao varrer janelas top-level: {e}")

    # 3) nada encontrado ainda -> devolve None para o wait_until continuar tentando
    print("[LOG] [Resumo] Ainda não consegui localizar a tela; retornando None.")
    return None



def _toggle_checkbox(parent_win: uia.Control, automation_id: str, name: str, desired: bool = True):
    cb = parent_win.CheckBoxControl(AutomationId=automation_id)
    if not cb.Exists(0, 0):
        cb = bfs_find(parent_win, name, types=("CheckBoxControl",), max_depth=10)
    if not cb:
        raise RuntimeError(f"Checkbox '{name}' não encontrada.")

    try:
        tp = cb.GetTogglePattern()
        if tp:
            state = tp.ToggleState  # 0=Off, 1=On, 2=Indeterminate
            if desired and state != 1:
                tp.Toggle()
            elif (not desired) and state == 1:
                tp.Toggle()
            return
    except Exception:
        pass

    uia_activate(cb, f"checkbox '{name}'")

def _find_edit_right_of_label(win: uia.Control, label_text: str) -> Optional[uia.Control]:
    label = bfs_find(win, label_text, types=("TextControl",), max_depth=12)
    if not label:
        return None
    lr = rect_of(label)
    if not lr:
        return None

    edits = []
    q = [(win, 0)]
    while q:
        node, d = q.pop(0)
        if d > 12:
            continue
        try:
            if node.ControlTypeName == "EditControl":
                rr = rect_of(node)
                if rr:
                    edits.append((rr, node))
            for ch in node.GetChildren():
                q.append((ch, d + 1))
        except Exception:
            pass

    best = None
    best_dx = None
    for rr, ed in edits:
        y_overlap = not (rr[3] < lr[1] or rr[1] > lr[3])
        if not y_overlap:
            continue
        if rr[0] <= lr[2]:
            continue
        dx = rr[0] - lr[2]
        if best is None or dx < best_dx:
            best = ed
            best_dx = dx
    return best

def gerar_resumo_folha_mensal():
    print("[LOG] Iniciando geração do Resumo da folha mensal...")
    win_pyget = ativar_e_maximizar()
    if not win_pyget:
        raise RuntimeError(f"Não consegui ativar a janela '{FOLHA_TITLE_EXATO}'.")

    main_win = ativar_e_maximizar()
    print(f"[LOG] Window UIA principal: {getattr(main_win, 'Name', '')!r}")
    try:
        main_win.SetFocus()
    except Exception:
        pass
    time.sleep(0.2)

    print("[LOG] Abrindo tela 'Resumo da folha mensal' via menus...")
    abrir_tela_resumo_folha()

    print("[LOG] Aguardando tela do Resumo aparecer (campo Mês/Ano por AutomationId)...")
  
    janela = wait_until(lambda: _find_resumo_container(main_win), timeout=60, interval=0.5)
    if not janela:
        print("[LOG] Falha: _find_resumo_container retornou None.")
        raise RuntimeError("Tela do Resumo não abriu (AutomationId do Mês/Ano não encontrado).")

    print(f"[LOG] Container do resumo detectado: {getattr(janela, 'Name', '')!r}")
    try:
        janela.SetFocus()
    except Exception:
        pass
    time.sleep(0.2)

    mes_ano = obter_mes_ano_anterior_slash()  # mantenha a função que você usa; nome de exemplo
    print(f"[LOG] [Resumo] Valor que será usado em Mês/Ano: {mes_ano}")

    # usamos a própria janela principal como container
    janela = ativar_e_maximizar() or main_win
    try:
        janela.SetFocus()
    except Exception:
        pass
    time.sleep(0.2)

    # =========== MÊS/ANO INICIAL / FINAL POR BOUNDINGRECTANGLE ===========
    print("[LOG] [Resumo] Procurando campo Mês/Ano INICIAL pela posição...")
    ed_mes_de = find_edit_by_rect(janela, RECT_MESANO_DE, "Mês/Ano INICIAL")

    print("[LOG] [Resumo] Procurando campo Mês/Ano FINAL pela posição...")
    ed_mes_ate = find_edit_by_rect(janela, RECT_MESANO_ATE, "Mês/Ano FINAL")

    if not ed_mes_de or not ed_mes_ate:
        print("[LOG] [Resumo] ERRO: não consegui localizar os campos Mês/Ano pela posição.")
        raise RuntimeError("Campos de Mês/Ano não localizados por BoundingRectangle.")

    print("[LOG] [Resumo] Preenchendo Mês/Ano INICIAL e FINAL...")
    set_text(ed_mes_de, mes_ano)
    set_text(ed_mes_ate, mes_ano)


    # ================== PROCESSAMENTO via AutomationId ==================
    print("[LOG] Procurando campo 'Processamento' pelo AutomationId...")
    ed_proc = janela.EditControl(AutomationId=EDIT_PROCESSAMENTO)
    if ed_proc.Exists(0, 0):
        rect_proc = rect_of(ed_proc)
        print(f"[LOG] Campo 'Processamento' localizado. Rect={rect_proc}, ID={EDIT_PROCESSAMENTO}")
        set_text(ed_proc, "2")
    else:
        print("[LOG] NÃO encontrei campo 'Processamento' por AutomationId; tentando por label...")
        ed_proc = _find_edit_right_of_label(janela, "Processamento")
        if ed_proc:
            print("[LOG] Campo 'Processamento' localizado por label; preenchendo com 2.")
            set_text(ed_proc, "2")
        else:
            print("[LOG] NÃO encontrei label/campo para 'Processamento' (vai ficar em branco).")

    # ================== CLASSIFICAÇÃO via AutomationId ==================
    print("[LOG] Procurando campo 'Classificação' pelo AutomationId...")
    ed_class = janela.EditControl(AutomationId=EDIT_CLASSIFICACAO)
    if ed_class.Exists(0, 0):
        rect_class = rect_of(ed_class)
        print(f"[LOG] Campo 'Classificação' localizado. Rect={rect_class}, ID={EDIT_CLASSIFICACAO}")
        set_text(ed_class, "1")
    else:
        print("[LOG] NÃO encontrei campo 'Classificação' por AutomationId; tentando por label...")
        ed_class = _find_edit_right_of_label(janela, "Classificação")
        if ed_class:
            print("[LOG] Campo 'Classificação' localizado por label; preenchendo com 1.")
            set_text(ed_class, "1")
        else:
            print("[LOG] NÃO encontrei label/campo para 'Classificação' (vai ficar em branco).")

    time.sleep(0.2)

    # ---- Checkboxes ----
    try:
        print("[LOG] Tentando marcar checkbox 'Pró-labore'...")
        _toggle_checkbox(janela, CHK_PROLABORE, "Listar recibos de pró-labore", True)
    except Exception as e:
        print(f"[WARN] Não consegui marcar pró-labore: {e}")

    try:
        print("[LOG] Tentando marcar checkbox 'Autônomos'...")
        _toggle_checkbox(janela, CHK_AUTONOMOS, "Listar recibos de autônomos", True)
    except Exception as e:
        print(f"[WARN] Não consegui marcar autônomos: {e}")

    try:
        print("[LOG] Tentando marcar checkbox 'Imprimir dados em destaque'...")
        _toggle_checkbox(janela, CHK_DESTAQUE, "Imprimir dados em destaque", True)
    except Exception as e:
        print(f"[WARN] Não consegui marcar destaque: {e}")

    # ---- Estabelecimentos ----    

    estab_de = find_edit_by_rect(janela, RECT_ESTAB_DE, "Estabelecimento INICIAL")
    estab_ate = find_edit_by_rect(janela, RECT_ESTAB_ATE, "Estabelecimento FINAL")


    # ---- Rádio "Geral" ----
    print("[LOG] Procurando radio button 'Geral'...")
    rb_geral = bfs_find(janela, "Geral", types=("RadioButtonControl",), max_depth=12)
    if rb_geral:
        print("[LOG] Detectei rádio 'Geral'; ativando.")
        uia_activate(rb_geral, "radio 'Geral'")
    else:
        print("[LOG] Não foi encontrado rádio 'Geral'.")

    # ---- Botão Imprimir ----
    print("[LOG] Procurando botão 'Imprimir' (AutomationId)...")
    btn_imprimir = janela.ButtonControl(AutomationId=BTN_IMPRIMIR_ID)
    if not btn_imprimir.Exists(0, 0):
        print("[LOG] Não achei por AutomationId; procurando por texto 'Imprimir'...")
        btn_imprimir = bfs_find(janela, "Imprimir", types=("ButtonControl",), max_depth=12)

    if not btn_imprimir:
        print("[LOG] NÃO encontrei botão 'Imprimir'.")
        raise RuntimeError("Botão 'Imprimir' não encontrado.")

    print("[LOG] Detectei botão 'Imprimir'; aguardando habilitar...")
    ok = wait_until(lambda: bool(getattr(btn_imprimir, "IsEnabled", True)), timeout=15, interval=0.3)
    if not ok:
        print("[WARN] Botão 'Imprimir' ainda desabilitado; tentando clicar mesmo assim.")

    uia_activate(btn_imprimir, "botão 'Imprimir'")
    time.sleep(0.8)
    print("[LOG] Clique em 'Imprimir' enviado.")


# ===================== Main =====================
def main():
    gerar_resumo_folha_mensal()


if __name__ == "__main__":
    main()
