# -*- coding: utf-8 -*-
"""
ERP Controle Patrimonial: exporta Demonstrativo Analítico mês a mês
"""

import time
from datetime import date
from dateutil.relativedelta import relativedelta
import uiautomation as uia
import pyautogui as pag
import pygetwindow as gw
from collections import deque
import calendar

# ============== CONFIG BÁSICA ==============
ERP_TITULO_PREFIX   = "Controle Patrimonial"    # título da janela principal
WORKSPACE_NAME      = "Espaço de trabalho"      # pane central
DEMONSTRATIVO       = "Demonstrativo Analítico"

DATA_CLICK_POS      = (725, 412)                # pixel do campo de data
HOTKEY_VISUALIZAR   = ('alt', 'i')              # Alt+I

# Toolbar (ajuste conforme seu viewer)
BOTAO_VIS_EXPORTAR_IDX   = 15                   # índice (1-based) do "Exportar"
BOTAO_VIS_APOS_SALVAR_IDX= 12                   # opcional

PASTA_DESTINO       = r"C:\Users\Usuario\OneDrive\Imagens\Scan"
INI, FIM            = date(2020,8,1), date(2023,8,1)

WAIT_SMALL, WAIT_MED, WAIT_LONG = 0.3, 0.8, 2.0
TIMEOUT_PADRAO      = 600

# pyautogui
pag.PAUSE = 0.05
pag.FAILSAFE = False

# ============== HELPERS GERAIS ==============
def _nome(ctrl):
    try: return (getattr(ctrl, "Name", "") or "").strip()
    except: return ""

def _tipo(ctrl):
    try: return getattr(ctrl, "ControlTypeName", "")
    except: return ""

def rect_of(ctrl):
    try:
        r = ctrl.BoundingRectangle
        return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except: return None

def wait_true(predicate, timeout=TIMEOUT_PADRAO, step=0.25):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if predicate():
                return True
        except: pass
        time.sleep(step)
    return False

def listar_botoes_toolbar(toolbar):
    """Lista botões (esq→dir) como no seu script: dedup por retângulo e ordena por X."""
    def _walk_buttons(node, acc):
        try:
            for c in node.GetChildren():
                if _tipo(c) == "ButtonControl":
                    acc.append(c)
                _walk_buttons(c, acc)
        except: pass

    bruto = []
    _walk_buttons(toolbar, bruto)
    seen, uniq = set(), []
    for b in bruto:
        rc = rect_of(b)
        if not rc: 
            continue
        if rc in seen:
            continue
        seen.add(rc)
        uniq.append((rc[0], rc, b))
    uniq.sort(key=lambda t: t[0])
    botoes = [b for _, _, b in uniq]
    print("\n[BARRA] Botões (esq→dir):")
    for i, b in enumerate(botoes, 1):
        print(f"  #{i:02d} {rect_of(b)} name='{_nome(b)}'")
    return botoes

def clicar_botao_por_indice(toolbar, indice_from_1):
    botoes = listar_botoes_toolbar(toolbar)
    if indice_from_1 is None:
        print("\n[DICA] Defina BOTAO_VIS_*_IDX para clicar automaticamente.")
        return False
    if indice_from_1 < 1 or indice_from_1 > len(botoes):
        print(f"[ERRO] Índice {indice_from_1} fora do intervalo 1..{len(botoes)}.")
        return False

    alvo = botoes[indice_from_1 - 1]

    # 1) InvokePattern
    try:
        inv = alvo.GetInvokePattern()
        if inv:
            inv.Invoke()
            print(f"[OK] Invoke no botão #{indice_from_1}.")
            return True
    except: pass

    # 2) LegacyIAccessible.DoDefaultAction
    try:
        leg = alvo.GetLegacyIAccessiblePattern()
        if leg:
            leg.DoDefaultAction()
            print(f"[OK] LegacyIAccessible default action no botão #{indice_from_1}.")
            return True
    except: pass

    # 3) Foco + Enter
    try:
        alvo.SetFocus()
        uia.SendKeys("{Enter}")
        print(f"[OK] Foco+Enter no botão #{indice_from_1}.")
        return True
    except: pass

    # 4) Clique físico no centro do botão (fallback final)
    try:
        rc = rect_of(alvo)
        if rc:
            cx = (rc[0] + rc[2]) // 2
            cy = (rc[1] + rc[3]) // 2
            pag.moveTo(cx, cy, duration=0.1)
            pag.click()
            print(f"[OK] Clique físico no botão #{indice_from_1}.")
            return True
    except: pass

    print(f"[ERRO] Falha ao acionar botão #{indice_from_1}.")
    return False


# ============== ATIVAÇÃO / LOCALIZAÇÃO ==============
def ativar_janela_principal():
    """Ativa/maximiza a maior janela cujo título começa com 'Controle Patrimonial'."""
    wins = [w for w in gw.getWindowsWithTitle(ERP_TITULO_PREFIX) if w.visible]
    if not wins:
        raise RuntimeError("Janela 'Controle Patrimonial' não localizada.")
    w = max(wins, key=lambda x: x.width * x.height)
    if not w.isActive: w.activate()
    if not w.isMaximized: w.maximize()
    time.sleep(1.0)
    return w

def localizar_demonstrativo():
    """Confere que a janela 'Demonstrativo Analítico' está aberta no workspace."""
    root = uia.GetRootControl()
    cont = None
    for w in root.GetChildren():
        if _tipo(w) == "WindowControl" and _nome(w).startswith(ERP_TITULO_PREFIX):
            cont = w; break
    if not cont:
        return None
    workspace = None
    for ch in cont.GetChildren():
        if _tipo(ch) == "PaneControl" and _nome(ch) == WORKSPACE_NAME:
            workspace = ch; break
    if not workspace:
        return None
    for ch in workspace.GetChildren():
        if _tipo(ch) == "WindowControl" and _nome(ch) == DEMONSTRATIVO:
            return ch
    return None

# ============== VISUALIZADOR + TOOLBAR (BASEADO NO SEU SCRIPT) ==============
def localizar_visualizador_e_toolbar():
    """
    Encontra o 'viewer' (WindowControl sem título) dentro do 'Espaço de trabalho' e,
    dentro dele, o Pane superior (20..50px de altura) como 'toolbar'.
    """
    root = uia.GetRootControl()

    # 1) janela Controle Patrimonial
    cont = None
    for w in root.GetChildren():
        if _tipo(w) == "WindowControl" and _nome(w).startswith(ERP_TITULO_PREFIX):
            cont = w; break
    if not cont:
        raise RuntimeError("Janela 'Controle Patrimonial' não localizada para o visualizador.")

    # 2) workspace
    workspace = None
    for ch in cont.GetChildren():
        if _tipo(ch) == "PaneControl" and _nome(ch) == WORKSPACE_NAME:
            workspace = ch; break
    if not workspace:
        raise RuntimeError("Pane 'Espaço de trabalho' não encontrado.")

    # 3) viewer: WindowControl com Name vazio
    viewer = None
    for ch in workspace.GetChildren():
        if _tipo(ch) == "WindowControl" and _nome(ch) == "":
            viewer = ch; break
    if not viewer:
        raise RuntimeError("Contêiner do visualizador não encontrado.")

    # 4) toolbar: Pane “raso” mais no topo
    toolbar, top_y = None, None
    for ch in viewer.GetChildren():
        if _tipo(ch) != "PaneControl": 
            continue
        r = rect_of(ch)
        if not r: 
            continue
        h = r[3] - r[1]
        if 20 <= h <= 50:  # mesma heurística do seu projeto
            if top_y is None or r[1] < top_y:
                top_y = r[1]
                toolbar = ch

    if not toolbar:
        raise RuntimeError("Barra de ferramentas do visualizador não encontrada.")

    try: viewer.SetFocus()
    except: pass
    return viewer, toolbar

# ============== FLUXO DE DATA/IMPRESSÃO ==============
def inserir_mes_erp(dt):
    """Clica no campo de data, digita MM/AAAA e depois DD_MM_AAAA (último dia do mês), Enter e abre visualização (Alt+I)."""
    mm_aaaa = f"{dt:%m/%Y}"
    ultimo_dia = calendar.monthrange(dt.year, dt.month)[1]
    dd_mm_aaaa = f"{ultimo_dia:02d}_{dt.month:02d}_{dt.year}"
    
    print(f"[DATA] Selecionando período {mm_aaaa} até {dd_mm_aaaa}")

    # garante foco na janela
    try:
        ativar_janela_principal()
    except: 
        pass

    # duplo clique no campo, seleciona tudo e digita
    x, y = DATA_CLICK_POS
    pag.moveTo(x, y, duration=0.1)
    pag.click(clicks=2, interval=0.08)
    time.sleep(WAIT_SMALL)

    # 1ª parte: último dia do mês
    pag.typewrite(dd_mm_aaaa, interval=0.02)
    time.sleep(WAIT_SMALL)
    pag.hotkey('Enter')

    # 2ª parte: mês/ano
    pag.hotkey('ctrl', 'a')
    pag.typewrite(mm_aaaa, interval=0.02)
    time.sleep(WAIT_SMALL)

    # confirma
    pag.press('enter')
    time.sleep(WAIT_MED)

    # abre visualização
    pag.hotkey(*HOTKEY_VISUALIZAR)
    time.sleep(WAIT_LONG)

# Aviso do Sistema GLOBAL (p.ex., "Salvar Relatório em PDF" → "Aviso do Sistema")
def _ctrl_type(c):
    try: return getattr(c, "ControlTypeName", "")
    except: return ""

def _value(c):
    try:
        vp = c.GetValuePattern()
        return (vp.Value or "").strip() if vp else ""
    except:
        return ""

def _parent(c):
    try: return c.GetParentControl()
    except: return None

def _collect_texts(ctrl):
    out = []
    try:
        if _ctrl_type(ctrl) in ("TextControl","EditControl"):
            nm = _nome(ctrl)
            if nm: out.append(nm)
        for ch in ctrl.GetChildren():
            out.extend(_collect_texts(ch))
    except: pass
    seen, res = set(), []
    for t in out:
        if t not in seen:
            seen.add(t); res.append(t)
    return res

def _dialog_ancestor(node):
    cur = node
    while cur:
        t = _ctrl_type(cur)
        if t in ("DialogControl","WindowControl"):
            return cur
        cur = _parent(cur)
    return node

def wait_global_aviso_do_sistema(timeout=30, interval=0.25, max_depth=8):
    root = uia.GetRootControl()
    end = time.time() + timeout

    def bfs_find_titlebar_or_dialog():
        q = deque([(root, 0)])
        while q:
            node, d = q.popleft()
            if d > max_depth:
                continue
            try:
                ctype = _ctrl_type(node)
                nm    = _nome(node)
                if ctype in ("DialogControl","WindowControl") and nm == "Aviso do Sistema":
                    return node, "\n".join(_collect_texts(node)) or nm
                if ctype == "TitleBarControl" and _value(node) == "Aviso do Sistema":
                    dlg = _dialog_ancestor(node)
                    return dlg, "\n".join(_collect_texts(dlg)) or "Aviso do Sistema"
            except:
                pass
            try:
                for ch in node.GetChildren():
                    q.append((ch, d+1))
            except:
                pass
        return None, None

    while time.time() < end:
        dlg, txt = bfs_find_titlebar_or_dialog()
        if dlg:
            return dlg, txt
        time.sleep(interval)
    return None, None


def salvar_como(pasta_destino, nome_arquivo_sem_ext):
    """
    Preenche o diálogo 'Salvar como':
      - Ctrl+L → caminho
      - Alt+N → nome 'demonstrativo mm.aaaa'
      - Enter (e só envia novamente se o diálogo NÃO fechar)
    """
    # Caixa de diálogo para caminho/confirmar
    time.sleep(1)
    pag.write(pasta_destino)
    time.sleep(1.0)
    pag.press('tab')
    time.sleep(0.3)
    pag.typewrite(nome_arquivo_sem_ext, interval=0.01)
    time.sleep(WAIT_MED)
    pag.press('tab'); pag.press('tab'); pag.press('space')
    pag.press('tab'); pag.press('tab'); pag.press('space')

    # Espera o "Aviso do Sistema" (global)
    dlg, dlg_text = wait_global_aviso_do_sistema(timeout=30, interval=0.25, max_depth=8)
    if dlg:
        print("\n[AVISO DO SISTEMA - relatório]")
        print(dlg_text if dlg_text else "(sem texto)")
        try: dlg.SetFocus()
        except: pass
        # aqui não apertamos Enter; o fluxo original usa Alt+S
    else:
        print("[INFO] Nenhum aviso do sistema detectado ao salvar relatório.")

    pag.hotkey('alt','s')
    time.sleep(0.5)
    pag.press('space')  # confirmar/fechar diálogo subsequente, se houver
    time.sleep(WAIT_MED)

# ============== MAIN LOOP ==============
def main():
    ativar_janela_principal()

    # garante que a tela do demonstrativo está aberta
    ok = wait_true(lambda: localizar_demonstrativo() is not None, timeout=15)
    if not ok:
        raise RuntimeError("Tela 'Demonstrativo Analítico' não localizada.")

    dt = INI
    while dt <= FIM:
        try:
            # 1) digita mês e abre visualização
            inserir_mes_erp(dt)

            # 2) aguarda e localiza viewer + toolbar
            ok = wait_true(lambda: localizar_visualizador_e_toolbar(), timeout=TIMEOUT_PADRAO)
            if not ok:
                raise RuntimeError("Visualizador não apareceu a tempo.")
            viewer, toolbar = localizar_visualizador_e_toolbar()

            # 3) clica Exportar
            listar_botoes_toolbar(toolbar)  # imprime para conferência
            if not clicar_botao_por_indice(toolbar, BOTAO_VIS_EXPORTAR_IDX):
                raise RuntimeError("Falha ao clicar 'Exportar' na toolbar.")

            # 4) Salvar como → pasta fixa + nome "demonstrativo mm.aaaa"
            nome = f"demonstrativo FILIAL {dt:%m_%Y}"
            salvar_como(PASTA_DESTINO, nome)

            # 5) RELOCALIZA a toolbar (alguns viewers a recriam após salvar)
            ok = wait_true(lambda: localizar_visualizador_e_toolbar(), timeout=10)
            if not ok:
                print("[WARN] Não consegui relocalizar o viewer/toolbar após salvar — tentando clicar na referência antiga.")
                # tenta na toolbar antiga mesmo; se falhar, o próximo mês ainda deve reabrir o viewer
                pass
            else:
                _, toolbar = localizar_visualizador_e_toolbar()

            # 6) OBRIGATÓRIO: clicar botão pós-salvar (fecha/prepara p/ próximo)
            if not clicar_botao_por_indice(toolbar, BOTAO_VIS_APOS_SALVAR_IDX):
                raise RuntimeError("Falha ao clicar botão pós-salvar na toolbar.")

            time.sleep(WAIT_MED)
            print(f"[OK] Mês {dt:%m/%Y} exportado como {nome}.")
        except Exception as e:
            print(f"[ERRO] Mês {dt:%m/%Y}: {e}")

        dt += relativedelta(months=1)
        time.sleep(WAIT_SMALL)

    print("\n==== FIM ====")

if __name__ == "__main__":
    main()
