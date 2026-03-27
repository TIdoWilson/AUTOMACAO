# -*- coding: utf-8 -*-
"""
ERP Contabilidade: exporta Balancete mês a mês (01/2020 .. 07/2025)
- Usa heurística de localização do visualizador/toolbar igual ao seu "importador final.py":
  * encontra o WindowControl do viewer vazio dentro de "Espaço de trabalho"
  * pega o Pane "raso" no topo (20..50px de altura) como toolbar
  * lista botões esq→dir e clica por índice
"""

import time
from datetime import date
from dateutil.relativedelta import relativedelta

import uiautomation as uia
import pyautogui as pag
import pygetwindow as gw
from collections import deque

# ============== CONFIG BÁSICA ==============
ERP_TITULO_PREFIX   = "Contabilidade"           # título da janela principal
WORKSPACE_NAME      = "Espaço de trabalho"      # pane central
BALANCETE_TITULO    = "Balancete de Verificação"

DATA_CLICK_POS      = (795, 388)                # pixel do campo de data
HOTKEY_VISUALIZAR   = ('alt', 'i')              # Alt+I

# Toolbar (ajuste conforme seu viewer)
BOTAO_VIS_EXPORTAR_IDX   = 11                   # índice (1-based) do "Exportar"
BOTAO_VIS_APOS_SALVAR_IDX= 13                   # opcional

PASTA_DESTINO       = r"C:\Users\Usuario\OneDrive\Imagens\Scan"
INI, FIM            = date(2024,1,1), date(2024,12,1)

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
    """Ativa/maximiza a maior janela cujo título começa com 'Contabilidade'."""
    wins = [w for w in gw.getWindowsWithTitle(ERP_TITULO_PREFIX) if w.visible]
    if not wins:
        raise RuntimeError("Janela 'Contabilidade' não localizada.")
    w = max(wins, key=lambda x: x.width * x.height)
    if not w.isActive: w.activate()
    if not w.isMaximized: w.maximize()
    time.sleep(1.0)
    return w

def localizar_balancete():
    """Confere que a janela 'Balancete de Verificação' está aberta no workspace."""
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
        if _tipo(ch) == "WindowControl" and _nome(ch) == BALANCETE_TITULO:
            return ch
    return None

# ============== VISUALIZADOR + TOOLBAR (BASEADO NO SEU SCRIPT) ==============
def localizar_visualizador_e_toolbar():
    """
    Encontra o 'viewer' (WindowControl sem título) dentro do 'Espaço de trabalho' e,
    dentro dele, o Pane superior (20..50px de altura) como 'toolbar'.
    """
    root = uia.GetRootControl()

    # 1) janela Contabilidade
    cont = None
    for w in root.GetChildren():
        if _tipo(w) == "WindowControl" and _nome(w).startswith(ERP_TITULO_PREFIX):
            cont = w; break
    if not cont:
        raise RuntimeError("Janela 'Contabilidade' não localizada para o visualizador.")

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
    """Usa pyautogui para garantir foco e digitação no campo de data, depois Alt+I."""
    mm_aaaa = f"{dt:%m/%Y}"
    print(f"[DATA] Selecionando {mm_aaaa}")

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
    pag.hotkey('ctrl', 'a')
    pag.typewrite(mm_aaaa, interval=0.02)
    time.sleep(WAIT_SMALL)
    pag.press('enter')
    time.sleep(WAIT_MED)

    # abre visualização
    pag.hotkey(*HOTKEY_VISUALIZAR)
    time.sleep(WAIT_LONG)

def salvar_como(pasta_destino, nome_arquivo_sem_ext):
    """
    Preenche o diálogo 'Salvar como':
      - Ctrl+L → caminho
      - Alt+N → nome 'balancete mm.aaaa'
      - Enter (e só envia novamente se o diálogo NÃO fechar)
    """
    # localiza o diálogo
    dlg = uia.WindowControl(ClassName='#32770')
    if not dlg.Exists(TIMEOUT_PADRAO):
        dlg = uia.WindowControl(NameRegex=r'(?i)salvar|salvar como')
        if not dlg.Exists(5):
            raise RuntimeError("Diálogo 'Salvar como' não apareceu.")

    try: dlg.SetFocus()
    except: pass
    time.sleep(WAIT_SMALL)

    # pasta
    pag.hotkey('ctrl', 'l')
    time.sleep(WAIT_SMALL)
    pag.typewrite(pasta_destino, interval=0.01)
    pag.press('enter')
    time.sleep(WAIT_MED)

    x, y = DATA_CLICK_POS
    pag.moveTo(x, y, duration=0.1)
    pag.click(clicks=2, interval=0.08)

    # nome arquivo
    pag.hotkey('alt', 'n')
    time.sleep(WAIT_SMALL)
    pag.hotkey('ctrl', 'a')
    pag.typewrite(nome_arquivo_sem_ext, interval=0.01)
    time.sleep(WAIT_SMALL)

    # salvar (1x)
    pag.press('enter')

    # aguarda o diálogo sumir; se ainda estiver aberto, confirma novamente
    def dlg_fechou():
        return not dlg.Exists(0.2)

    if not wait_true(dlg_fechou, timeout=5, step=0.2):
        # pode ser prompt de confirmação/sobrescrita
        pag.press('enter')
        wait_true(dlg_fechou, timeout=5, step=0.2)

    time.sleep(WAIT_MED)

# ============== MAIN LOOP ==============
def main():
    ativar_janela_principal()

    # garante que a tela do Balancete está aberta
    ok = wait_true(lambda: localizar_balancete() is not None, timeout=15)
    if not ok:
        raise RuntimeError("Tela 'Balancete de Verificação' não localizada.")

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

            # 4) Salvar como → pasta fixa + nome "balancete mm.aaaa"
            nome = f"balancete {dt:%m_%Y}"
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
