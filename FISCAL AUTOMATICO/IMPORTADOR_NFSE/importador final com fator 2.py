# -*- coding: utf-8 -*-
import os
import re
import time
import datetime as dt
import xml.etree.ElementTree as ET
from collections import deque

import numpy as np
import pandas as pd
import pyautogui as pag
import pygetwindow as gw
import uiautomation as uia
import unicodedata

# ===================== CONFIG =====================
CAMINHO_PASTA_BASE = r'\\192.0.0.251\arquivos\XML PREFEITURA'

# Tentativas de caminho para a planilha LISTA.xlsx (usa o primeiro que existir)
CAMINHOS_PLANILHA_POSSIVEIS = [
    r'\\192.0.0.251\arquivos\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\IMPORTADOR_NFSE\LISTA EMPRESAS SISTEMA.XLSX',
]
try:
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    CAMINHOS_PLANILHA_POSSIVEIS.insert(1, os.path.join(_script_dir, 'LISTA.xlsx'))
except Exception:
    pass

CAMINHO_LOG = r"\\192.0.0.251\arquivos\XML PREFEITURA\log_importacao.txt"

# Palavras-chave de UI / títulos
FISCAL_TITLE_EXATO = "Fiscal"
WORKSPACE_NAME     = "Espaço de trabalho"
NFSE_TITULO_PREFIX = "Importação de Nota Fiscal de Serviço Eletrônica"
BTN_CARREGAR       = "carregar"
BTN_IMPORTAR       = "importar"
BTN_APURACAO       = "apuracao"
BTN_GRAVAR         = "gravar"


MENU_JANELA = "Janela"
SUB_FECHAR_TODAS = "Fechar Todas"

# Relatórios > Legais > Livro Registro de ISS > Serviços Prestados - Padrão
MENU_RELATORIOS           = "relatórios"
SUB_LEGAIS                = "legais"
ITEM_LIVRO_REG_ISS        = "livro registro de iss"
ITEM_SERV_PREST_PADRAO    = "serviços prestados - padrão"
JANELA_LIVRO              = "Livro Registro de Serviços Prestados - Padrão"
BTN_IMPRIMIR              = "imprimir"

# Tributos > Simples Nacional > Valor Folha - Fator R
MENU_TRIBUTOS             = "Tributos"
SUB_SIMPLES_NACIONAL      = "Simples Nacional"
MENU_FATOR_R              = "Valor Folha - Fator R"
ABA_CPP                   = "Pesquisa de informar anexos busca CPP"  

# Toolbar do visualizador — índices (esquerda→direita).
# Deixe None para apenas listar e você confirmar o número no console.
BOTAO_VIS_EXPORTAR_IDX = 16
BOTAO_VIS_APOS_SALVAR_IDX = 13

# pyautogui
pag.PAUSE = 0.05
pag.FAILSAFE = True

# ======= CONFIG da detecção de campos (topo-direito) =======
ROI_RIGHT_WIDTH = 560
ROI_TOP_OFFSET  = 34
ROI_HEIGHT      = 130
GRID_COLS       = 10
GRID_ROWS       = 6
MIN_W, MAX_W    = 18, 160
MIN_H, MAX_H    = 14, 36

# ===================== UTIL (datas, strings) =====================
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

def obter_mes_anterior_pasta() -> str:
    mes_anterior = pd.Timestamp.today().to_period('M') - 1
    return f"{int(mes_anterior.month):02d}.{mes_anterior.year}"

def _strip_accents(s: str) -> str:
    s = s or ""
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")

def _canon(s: str) -> str:
    return " ".join(_strip_accents(s).lower().split())

def _contains_canon(sub: str, text: str) -> bool:
    return _canon(sub) in _canon(text)

# ===================== UTIL (arquivos/planilha/xml/log) =====================
def _primeiro_caminho_existente(caminhos: list[str]) -> str | None:
    for c in caminhos:
        try:
            if os.path.exists(c):
                return c
        except Exception:
            continue
    return None

def _normaliza_head(cols):
    return [str(c).strip().lower() for c in cols]

def _to_str(x):
    if x is None:
        return ""
    if isinstance(x, float) and np.isnan(x):
        return ""
    s = str(x)
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s.strip()

def carregar_mapa_por_cnpj() -> dict:
    caminho = _primeiro_caminho_existente(CAMINHOS_PLANILHA_POSSIVEIS)
    if not caminho:
        raise FileNotFoundError("Planilha LISTA.xlsx não encontrada.")
    df = pd.read_excel(caminho, dtype=object, engine="openpyxl")
    df.columns = _normaliza_head(df.columns)

    # localizar colunas principais
    col_cnpj   = next((c for c in df.columns if c.replace('ç','c').startswith('cnpj')), None)
    col_codigo = next((c for c in df.columns if 'codigo' in c and 'erp' in c), None) \
                 or next((c for c in df.columns if c in ('codigo_erp','código_erp','cod_erp','codigoserp')), None)
    col_estab  = next((c for c in df.columns if 'estab' in c), None) or 'estabelecimento'

    # coluna opcional: fator (p.ex. 'fator', 'fator r', 'valor_fator')
    col_fator  = next((c for c in df.columns if 'fator' in c), None)

    if not col_cnpj or not col_codigo or not col_estab:
        raise KeyError(f"Colunas não encontradas. Cabeçalho: {df.columns.tolist()}")

    mapa = {}
    for _, row in df.iterrows():
        cnpj_digits = re.sub(r'\D', '', _to_str(row.get(col_cnpj, "")))
        if not cnpj_digits:
            continue
        codigo_erp = _to_str(row.get(col_codigo, ""))
        estab  = _to_str(row.get(col_estab, "")) or "1"
        fatorR  = _to_str(row.get(col_fator, "")) if col_fator else ""
        if codigo_erp:
            mapa[cnpj_digits] = {
                'codigo_erp': codigo_erp,
                'estabelecimento': estab,
                'fator': fatorR
            }
    return mapa

def encontrar_xmls(base_dir: str) -> list[str]:
    alvo = obter_mes_anterior_pasta()
    out = []
    for root, _, files in os.walk(base_dir):
        if not root.endswith(alvo):
            continue
        for f in files:
            if f.lower().endswith(".xml"):
                out.append(os.path.join(root, f))
    return out

def extrair_cnpj(xml_path: str) -> str | None:
    try:
        ns = {'nfs': 'https://www.esnfs.com.br/xsd'}
        tree = ET.parse(xml_path)
        root = tree.getroot()
        nfs_node = root.find('.//nfs:nfs', ns)
        if nfs_node is None:
            return None

        if xml_path.lower().endswith('emitido.xml'):
            prestador = nfs_node.find('nfs:prestadorServico', ns)
            if prestador is not None:
                return _to_str(prestador.find('nfs:nrDocumento', ns).text)
        elif xml_path.lower().endswith('recebido.xml'):
            tomador = nfs_node.find('nfs:tomadorServico', ns)
            if tomador is not None:
                return _to_str(tomador.find('nfs:nrDocumento', ns).text)
    except Exception:
        return None
    return None

def montar_log_empresas() -> list[str]:
    linhas = []
    mapa = carregar_mapa_por_cnpj()
    pastas_ok = set()
    for xml in encontrar_xmls(CAMINHO_PASTA_BASE):
        pasta = os.path.dirname(xml)
        if pasta in pastas_ok:
            continue
        cnpj = extrair_cnpj(xml)
        if not cnpj:
            linhas.append(f"[ERRO] Não foi possível extrair CNPJ do arquivo: {xml}")
            continue
        cnpj_digits = re.sub(r'\D', '', _to_str(cnpj))
        info = mapa.get(cnpj_digits)
        if not info:
            linhas.append(f"[ERRO] CNPJ {cnpj_digits} não encontrado na planilha - Pasta: {pasta}")
            continue
        pastas_ok.add(pasta)
        codigo = _to_str(info.get('codigo_erp', ''))
        estab  = _to_str(info.get('estabelecimento', '1')) or '1'
        fatorR = _to_str(info.get('fator', '')) or '-'
        linhas.append(f"[OK] Código {codigo} Estab {estab} Fator {fatorR} - Pasta: {pasta}") 
    return linhas

def registrar_log(linhas_log: list[str]) -> None:
    with open(CAMINHO_LOG, 'w', encoding='utf-8') as f:
        f.write(f'Log gerado em {dt.datetime.now():%d/%m/%Y %H:%M:%S}\n')
        f.write('='*60 + '\n')
        f.writelines(l + '\n' for l in linhas_log)
        f.write('='*60 + '\n')

def carregar_resultados_log(caminho_log: str) -> list[tuple[str, str, str, str]]:
    out = []
    if not os.path.exists(caminho_log):
        return out
    with open(caminho_log, 'r', encoding='utf-8') as f:
        for linha in f:
            if not linha.startswith('[OK]'):
                continue

            m = re.search(
                r"^\[OK\]\s+Código\s+(?P<cod>\S+)\s+Estab\s+(?P<estab>\S+)\s+Fator\s+(?P<fatorR>.*?)\s+-\s+Pasta:\s+(?P<pasta>.+)$",
                linha.strip()
            )
            if m:
                out.append((m.group('cod'), m.group('estab'), m.group('fatorR'), m.group('pasta')))
    return out


def uia_activate_fast(ctrl, name_for_log="(controle)") -> bool:
    """Ativa rapidamente sem esperas excessivas."""
    if not ctrl:
        return False
    try:
        ip = ctrl.GetInvokePattern()
        if ip:
            ip.Invoke()
            time.sleep(0.1)
            return True
    except Exception:
        pass
    try:
        bring_into_view(ctrl)
        ctrl.SetFocus()
        time.sleep(0.1)
        return True
    except Exception:
        pass
    return False

# ===================== UIA Helpers =====================
def normalize(s):
    try: return (s or "").strip().casefold()
    except: return ""

def wait_until(fn, timeout=20.0, interval=0.25, on_wait=None):
    end = time.time() + timeout
    while time.time() < end:
        try:
            v = fn()
            if v:
                return v
        except:
            pass
        if on_wait:
            try: on_wait()
            except: pass
        time.sleep(interval)
    return None

def bfs_find(root_ctrl, name_substr,
             types=('WindowControl','PaneControl','GroupControl','DocumentControl','ButtonControl','EditControl','MenuItemControl','ListItemControl','TreeItemControl','TabItemControl'),
             max_depth=6):
    target = normalize(name_substr)
    q = deque([(root_ctrl, 0)])
    while q:
        node, depth = q.popleft()
        if depth > max_depth:
            continue
        try:
            if target and target in normalize(getattr(node, 'Name', None)):
                if not types or node.ControlTypeName in types:
                    return node
            for child in node.GetChildren():
                q.append((child, depth + 1))
        except:
            pass
    return None

def bring_into_view(ctrl):
    try:
        sip = ctrl.GetScrollItemPattern()
        if sip:
            sip.ScrollIntoView()
    except:
        pass
    try:
        ctrl.SetFocus()
    except:
        pass

def uia_activate(ctrl, name_for_log="(controle)", prefer_invoke=True):
    try:
        bring_into_view(ctrl)
        if prefer_invoke:
            try:
                inv = ctrl.GetInvokePattern()
                if inv:
                    inv.Invoke()
                    print(f"[OK] Invoke → {name_for_log}")
                    return True
            except:
                pass
        try:
            sel = ctrl.GetSelectionItemPattern()
            if sel:
                sel.Select()
                print(f"[OK] SelectionItem.Select → {name_for_log}")
                return True
        except:
            pass
        try:
            ec = ctrl.GetExpandCollapsePattern()
            if ec:
                try:
                    ec.Expand()
                except:
                    pass
                print(f"[OK] ExpandCollapse.Expand → {name_for_log}")
                return True
        except:
            pass
    except Exception as e:
        print(f"[ERRO] uia_activate({name_for_log}): {e}")
    print(f"[FALHA] Não foi possível ativar {name_for_log} por UIA.")
    return False

def find_first_by_subname(scopes, subname, types, max_depth=6):
    for scope in scopes:
        c = bfs_find(scope, subname, types=types, max_depth=max_depth)
        if c:
            return c
    return None

def rect_of(ctrl):
    try:
        r = ctrl.BoundingRectangle
        return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        return None

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

def unique_add(lst, seen_rects, ctrl):
    r = rect_of(ctrl)
    if not r or r in seen_rects:
        return
    seen_rects.add(r)
    lst.append(ctrl)

def read_value(ctrl):
    try:
        return ctrl.GetValuePattern().Value or ""
    except Exception:
        return ""

def normalize_digits(s):
    return re.sub(r'\D', '', s or '')

def is_date_like(s):
    s = s or ""
    if "/" in s and re.search(r"\d{1,2}/\d{1,2}/\d{2,4}", s):
        return True
    if len(normalize_digits(s)) >= 8:
        return True
    return False

# ===================== ERP: ativar / troca empresa =====================
def ativar_e_maximizar():
    wins = [w for w in gw.getWindowsWithTitle(FISCAL_TITLE_EXATO) if w.visible]
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

def set_text(control: uia.Control, text: str):
    try:
        vp = control.GetValuePattern()
        vp.SetValue(text)
        return
    except Exception:
        pass
    try:
        control.SetFocus()        # foca sem mover o mouse
        uia.SendKeys('^a{DEL}')
        uia.SendKeys(text)
    except Exception as e:
        raise RuntimeError(f'Não foi possível escrever no controle {control}: {e}')

def trocar_empresa(win_window, codigo: str, estabelecimento: str, data_ddmmaa: str):
    if not win_window:
        raise RuntimeError("Janela do Fiscal não está ativa.")

    L = win_window.left + max(0, win_window.width - ROI_RIGHT_WIDTH)
    T = win_window.top + ROI_TOP_OFFSET
    R = win_window.left + win_window.width - 8
    B = T + ROI_HEIGHT

    xs = [int(L + (R - L) * (i + 0.5) / GRID_COLS) for i in range(GRID_COLS)]
    ys = [int(T + (B - T) * (j + 0.5) / GRID_ROWS) for j in range(GRID_ROWS)]

    candidatos, seen = [], set()
    for y in ys:
        for x in xs:
            try:
                el = uia.ControlFromPoint(x, y)
                for _ in range(5):
                    if not el:
                        break
                    if is_edit_candidate(el):
                        unique_add(candidatos, seen, el)
                        break
                    el = el.GetParentControl()
            except Exception:
                continue

    if not candidatos:
        raise RuntimeError("Não encontrei campos na ROI. Ajuste ROI_*.")

    filtrados, seen2 = [], set()
    for c in candidatos:
        if is_edit_candidate(c):
            unique_add(filtrados, seen2, c)

    filtrados.sort(key=lambda c: (rect_of(c)[1], rect_of(c)[0]))
    if len(filtrados) < 2:
        raise RuntimeError("Poucos campos encontrados para empresa/estab/data.")

    campo_data = None
    restantes = []
    for c in filtrados:
        v = read_value(c)
        if is_date_like(v):
            campo_data = c
        else:
            restantes.append(c)
    if not campo_data:
        mesma_linha = [c for c in filtrados if abs(rect_of(c)[1]-rect_of(filtrados[0])[1]) <= 10]
        campo_data = sorted(mesma_linha, key=lambda c: rect_of(c)[0])[-1]

    if len(restantes) < 2:
        restantes = [c for c in filtrados if c is not campo_data]
    restantes.sort(key=lambda c: (rect_of(c)[1], rect_of(c)[0]))
    campo_codigo, campo_estab = restantes[0], (restantes[1] if len(restantes) > 1 else None)

    # preencher + ENTERs entre empresa e estab
    set_text(campo_codigo, codigo)
    try: campo_codigo.SetFocus()
    except: pass
    uia.SendKeys('{Enter}')
    uia.SendKeys('{Enter}')
    time.sleep(1.0)

    if campo_estab:
        set_text(campo_estab, estabelecimento)
    set_text(campo_data, data_ddmmaa)
    uia.SendKeys('{Enter}')
    uia.SendKeys('{Enter}')

# ===================== NFS-e: localizar janela/botões/avisos =====================
def _nome(ctrl):
    try: return (getattr(ctrl, "Name", "") or "").strip()
    except: return ""

def _tipo(ctrl):
    try: return getattr(ctrl, "ControlTypeName", "")
    except: return ""

def localizar_nfse_dentro_fiscal():
    root = uia.GetRootControl()

    # 1) Fiscal (WindowControl de 1º nível)
    fiscal = None
    for w in root.GetChildren():
        try:
            if _tipo(w) != "WindowControl":
                continue
            if (_nome(w)).strip().startswith(FISCAL_TITLE_EXATO):
                fiscal = w
                break
        except:
            continue
    if not fiscal:
        raise RuntimeError("Janela 'Fiscal' não localizada no Desktop UIA.")

    # 2) Pane 'Espaço de trabalho'
    workspace = None
    for ch in fiscal.GetChildren():
        try:
            if _tipo(ch) == "PaneControl" and _nome(ch) == WORKSPACE_NAME:
                workspace = ch
                break
        except:
            continue
    if not workspace:
        raise RuntimeError("Pane 'Espaço de trabalho' não encontrado dentro do Fiscal.")

    # 3) Janela de NFS-e por PREFIXO (aceita qualquer município/UF após)
    janela_nfse = None
    for ch in workspace.GetChildren():
        try:
            if _tipo(ch) == "WindowControl" and _nome(ch).startswith(NFSE_TITULO_PREFIX):
                janela_nfse = ch
                break
        except:
            continue
    if not janela_nfse:
        raise RuntimeError(
            f"Janela que começa com '{NFSE_TITULO_PREFIX}' não encontrada dentro de '{WORKSPACE_NAME}'."
        )

    try:
        janela_nfse.SetFocus()
    except:
        pass
    return janela_nfse

def localizar_fator_dentro_fiscal():
    root = uia.GetRootControl()

    # 1) Fiscal (WindowControl de 1º nível)
    fiscal = None
    for w in root.GetChildren():
        try:
            if _tipo(w) != "WindowControl":
                continue
            if (_nome(w)).strip().startswith(FISCAL_TITLE_EXATO):
                fiscal = w
                break
        except:
            continue
    if not fiscal:
        raise RuntimeError("Janela 'Fiscal' não localizada no Desktop UIA.")

    # 2) Pane 'Espaço de trabalho'
    workspace = None
    for ch in fiscal.GetChildren():
        try:
            if _tipo(ch) == "PaneControl" and _nome(ch) == WORKSPACE_NAME:
                workspace = ch
                break
        except:
            continue
    if not workspace:
        raise RuntimeError("Pane 'Espaço de trabalho' não encontrado dentro do Fiscal.")

    # 3) Janela da fator R
    janela_fator = None
    for ch in workspace.GetChildren():
        try:
            if _tipo(ch) == "WindowControl" and _nome(ch).startswith(MENU_FATOR_R):
                janela_fator = ch
                break
        except:
            continue
    if not janela_fator:
        raise RuntimeError(
            f"Janela que começa com '{MENU_FATOR_R}' não encontrada dentro de '{WORKSPACE_NAME}'."
        )

    try:
        janela_fator.SetFocus()
    except:
        pass
    return janela_fator

def localizar_cpp_dentro_fiscal():
    root = uia.GetRootControl()

    # 1) Fiscal (WindowControl de 1º nível)
    fiscal = None
    for w in root.GetChildren():
        try:
            if _tipo(w) != "WindowControl":
                continue
            if (_nome(w)).strip().startswith(FISCAL_TITLE_EXATO):
                fiscal = w
                break
        except:
            continue
    if not fiscal:
        raise RuntimeError("Janela 'Fiscal' não localizada no Desktop UIA.")

    # 2) Pane 'Espaço de trabalho'
    workspace = None
    for ch in fiscal.GetChildren():
        try:
            if _tipo(ch) == "PaneControl" and _nome(ch) == WORKSPACE_NAME:
                workspace = ch
                break
        except:
            continue
    if not workspace:
        raise RuntimeError("Pane 'Espaço de trabalho' não encontrado dentro do Fiscal.")

    # 3) Janela da CPP
    janela_cpp = None
    for ch in workspace.GetChildren():
        try:
            if _tipo(ch) == "WindowControl" and _nome(ch).startswith(ABA_CPP):
                janela_cpp = ch
                break
        except:
            continue
    if not janela_cpp:
        raise RuntimeError(
            f"Janela que começa com '{ABA_CPP}' não encontrada dentro de '{WORKSPACE_NAME}'."
        )

    try:
        janela_cpp.SetFocus()
    except:
        pass
    return janela_cpp

def encontrar_botoes_nfse(janela_nfse):
    btn_carregar = bfs_find(janela_nfse, BTN_CARREGAR, types=('ButtonControl',), max_depth=8)
    btn_importar = bfs_find(janela_nfse, BTN_IMPORTAR, types=('ButtonControl',), max_depth=8)
    return btn_carregar, btn_importar

def encontrar_botoes_fator(janela_fator):
    btn_apuracao = bfs_find(janela_fator, BTN_APURACAO, types=('ButtonControl',), max_depth=8)
    btn_carregar = bfs_find(janela_fator, BTN_CARREGAR, types=('ButtonControl',), max_depth=8)
    btn_gravar = bfs_find(janela_fator, BTN_GRAVAR, types=('ButtonControl',), max_depth=8)
    return btn_apuracao, btn_carregar, btn_gravar

def _iter_buttons(ctrl):
    try:
        for ch in ctrl.GetChildren():
            if getattr(ch, "ControlTypeName", "") == "ButtonControl":
                yield ch
            yield from _iter_buttons(ch)
    except Exception:
        return

def focus_and_dismiss_alert(alert_win):
    try: alert_win.SetFocus()
    except: pass
    time.sleep(0.05)
    preferidos = {"ok", "sim", "yes", "fechar", "confirmar", "entendi"}
    try:
        for btn in _iter_buttons(alert_win):
            nm = (_nome(btn)).lower()
            if any(p in nm for p in preferidos):
                try:
                    inv = btn.GetInvokePattern()
                    if inv:
                        inv.Invoke()
                        return True
                except:
                    try: btn.SetFocus()
                    except: pass
                    uia.SendKeys('{Enter}')
                    return True
        for btn in _iter_buttons(alert_win):
            try:
                inv = btn.GetInvokePattern()
                if inv:
                    inv.Invoke()
                    return True
            except:
                try: btn.SetFocus()
                except: pass
                uia.SendKeys('{Enter}')
                return True
    except:
        pass
    uia.SendKeys('{Enter}')
    return True

def focus_and_dismiss_alert_gdfe(alert_win):
    """
    Versão específica para o aviso da GDFe.
    Trabalha apenas com a janela recebida (alert_win),
    para não interferir nos outros fluxos de aviso.
    """
    try:
        alert_win.SetFocus()
    except:
        pass
    time.sleep(1)

    preferidos = {"ok"}

    try:
        # Tenta primeiro botões com textos "OK", "Sim", etc.
        for btn in _iter_buttons(alert_win):
            nm = (_nome(btn)).lower()
            if any(p in nm for p in preferidos):
                try:
                    inv = btn.GetInvokePattern()
                    if inv:
                        inv.Invoke()
                        return True
                except:
                    try:
                        btn.SetFocus()
                    except:
                        pass
                    uia.SendKeys('{Enter}')
                    return True

        # Se não achou pelos nomes acima, tenta o primeiro botão que aceitar Invoke/Enter
        for btn in _iter_buttons(alert_win):
            try:
                inv = btn.GetInvokePattern()
                if inv:
                    inv.Invoke()
                    return True
            except:
                try:
                    btn.SetFocus()
                except:
                    pass
                uia.SendKeys('{Enter}')
                return True
    except:
        pass

    # Fallback final – Enter “cego”
    uia.SendKeys('{Enter}')
    return True

def _fiscal_window():
    root = uia.GetRootControl()
    for w in root.GetChildren():
        if _tipo(w) == "WindowControl" and _nome(w).startswith("Fiscal"):
            return w
    return None

def get_fiscal_from_nfse(janela_nfse):
    p = janela_nfse
    while p:
        if _tipo(p) == "WindowControl" and _nome(p).startswith("Fiscal"):
            return p
        try: p = p.GetParentControl()
        except: break
    return _fiscal_window()

def _collect_text(ctrl):
    txts = []
    try:
        if _tipo(ctrl) in ("TextControl","EditControl"):
            nm = _nome(ctrl)
            if nm: txts.append(nm)
        for ch in ctrl.GetChildren():
            txts.extend(_collect_text(ch))
    except: pass
    seen, out = set(), []
    for t in txts:
        if t not in seen:
            seen.add(t); out.append(t)
    return out

def wait_aviso_do_sistema(fiscal_win, timeout=600, interval=0.25):
    end = time.time() + timeout
    while time.time() < end:
        try:
            for ch in fiscal_win.GetChildren():
                if _tipo(ch) == "WindowControl" and _nome(ch) == "Aviso do Sistema":
                    texto = "\n".join(_collect_text(ch)) or _nome(ch)
                    return ch, texto
        except: pass
        time.sleep(interval)
    return None, None

def fechar_aviso_do_sistema(ctrl_aviso):
    try: ctrl_aviso.SetFocus()
    except: pass
    time.sleep(0.05)
    try:
        for btn in ctrl_aviso.GetChildren():
            if _tipo(btn) == "ButtonControl" and _nome(btn).lower() == "ok":
                try:
                    inv = btn.GetInvokePattern()
                    if inv:
                        inv.Invoke()
                        return True
                except:
                    try: btn.SetFocus()
                    except: pass
                    uia.SendKeys("{Enter}")
                    return True
    except: pass
    uia.SendKeys("{Enter}")
    return True

def fechar_todas_as_janelas():
    """Fecha todas as janelas abertas no ERP."""
    root = uia.GetRootControl()
    workspace = bfs_find(root, WORKSPACE_NAME,
                        types=("PaneControl", "GroupControl", "DocumentControl"),
                        max_depth=4) or root

    jan = wait_until(
        lambda: find_first_by_subname([root, workspace], MENU_JANELA, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if jan:
        uia_activate_fast(jan)
        time.sleep(0.2)
    else:
        print("[AVISO] Menu 'Janela' não encontrado")
        return

    janela = wait_until(
        lambda: find_first_by_subname([root, workspace], SUB_FECHAR_TODAS, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if janela:
        uia_activate_fast(janela)
        time.sleep(0.2)
    else:
        print("[AVISO] Submenu 'Fechar Todas' não encontrado")
        return
    
    dlg, dlg_text = wait_global_aviso_do_sistema(timeout=30, interval=0.25, max_depth=8)
    if dlg:
        print("\n[AVISO DO SISTEMA - relatório]")
        print(dlg_text if dlg_text else "(sem texto)")
        try: dlg.SetFocus()
        except: pass

    time.sleep(0.2)
    pag.hotkey('alt','s')

def tentar_importar_com_retry_gdfe(btn_importar_enabled, fiscal_win, max_tentativas=3):
    """
    Clica no botão Importar e trata especificamente o aviso da GDFe:

      "Não foi possível buscar dados iniciais da GDFe. Tente novamente mais tarde."

    Regras:
    - Se existir QUALQUER 'Aviso do Sistema' pendente, ele é SEMPRE tratado antes
      de tentar acionar o botão Importar.
    - Se o aviso for da GDFe, fecha com focus_and_dismiss_alert_gdfe e tenta
      novamente até max_tentativas.
    - Para outros avisos, usa focus_and_dismiss_alert normal e considera o fluxo
      encerrado (uma vez só).
    - Retorna True se considerar o IMPORTAR concluído, False se desistir.
    """
    if not fiscal_win:
        # Sem referência à janela do Fiscal, não temos como detectar/fechar avisos
        print("[WARN] fiscal_win não definido em tentar_importar_com_retry_gdfe.")
        return False

    chave_gdfe = "buscar dados iniciais da gdfe"

    for tentativa in range(1, max_tentativas + 1):
        print(f"\n>>> Acionando 'Importar' (tentativa {tentativa}/{max_tentativas})...")

        # 1) PRIMEIRO: tratar qualquer Aviso do Sistema já aberto
        aviso_existente, txt_existente = wait_aviso_do_sistema(
            fiscal_win, timeout=0.5, interval=0.1
        )
        if aviso_existente:
            msg = (txt_existente or "").strip()
            msg_low = msg.lower()
            print("\n[AVISO DO SISTEMA - pendente antes de IMPORTAR]")
            print(msg if msg else "(sem texto)")

            if chave_gdfe in msg_low:
                # Aviso de GDFe deve SEMPRE ser tratado antes de qualquer clique
                print("[INFO] Aviso de GDFe detectado (antes de clicar em Importar). "
                      "Fechando e tentando novamente...")
                focus_and_dismiss_alert_gdfe(aviso_existente)
                time.sleep(0.5)

                if tentativa < max_tentativas:
                    # volta para o loop e tenta de novo (começando de novo pela verificação de aviso)
                    continue
                else:
                    print("[WARN] Limite de tentativas para erro GDFe atingido "
                          "(aviso persistente antes de clicar em Importar).")
                    return False
            else:
                # Outros avisos são tratados com o handler padrão e encerram o fluxo
                focus_and_dismiss_alert(aviso_existente)
                time.sleep(0.3)
                return True

        # 2) Agora tentamos acionar o botão Importar
        ok = uia_activate(btn_importar_enabled, "botão 'Importar'")
        if not ok:
            print("[ERRO] Falha ao acionar botão 'Importar' por UIA.")
            return False

        # 3) Depois de clicar em Importar, esperamos um possível Aviso do Sistema
        aviso2, txt2 = wait_aviso_do_sistema(fiscal_win, timeout=600, interval=0.25)

        if not aviso2:
            # Nenhum aviso → consideramos que a importação prosseguiu normalmente
            print("[INFO] Nenhum 'Aviso do Sistema' após IMPORTAR.")
            return True

        msg = (txt2 or "").strip()
        msg_low = msg.lower()
        print("\n[AVISO DO SISTEMA - após IMPORTAR]")
        print(msg if msg else "(sem texto)")

        if chave_gdfe in msg_low:
            # Aviso específico da GDFe após IMPORTAR
            print("[INFO] Aviso de GDFe detectado (após IMPORTAR). "
                  "Fechando e tentando importar novamente...")
            focus_and_dismiss_alert_gdfe(aviso2)
            time.sleep(0.5)

            if tentativa < max_tentativas:
                # volta para o loop e recomeça (sempre tratando o aviso primeiro)
                continue
            else:
                print("[WARN] Limite de tentativas para erro GDFe atingido "
                      "(após IMPORTAR).")
                return False
        else:
            # Aviso normal: usa o handler padrão e encerra
            focus_and_dismiss_alert(aviso2)
            time.sleep(0.3)
            return True

    return False

# ===================== Visualizador/Relatório =====================
def _find_fiscal_root():
    root = uia.GetRootControl()
    for w in root.GetChildren():
        if _tipo(w) == "WindowControl" and _nome(w).startswith("Fiscal"):
            return w
    return None

def localizar_visualizador_relatorio():
    fiscal = _find_fiscal_root()
    if not fiscal:
        raise RuntimeError("Janela 'Fiscal' não localizada para o visualizador.")
    workspace = None
    for ch in fiscal.GetChildren():
        if _tipo(ch) == "PaneControl" and _nome(ch) == WORKSPACE_NAME:
            workspace = ch
            break
    if not workspace:
        raise RuntimeError("Pane 'Espaço de trabalho' não encontrado.")
    viewer = None
    for ch in workspace.GetChildren():
        if _tipo(ch) == "WindowControl" and _nome(ch) == "":
            viewer = ch
            break
    if not viewer:
        raise RuntimeError("Contêiner do visualizador não encontrado.")

    # barra de ferramentas: Pane baixo e perto do topo
    toolbar, top_y = None, None
    for ch in viewer.GetChildren():
        if _tipo(ch) != "PaneControl":
            continue
        r = rect_of(ch)
        if not r:
            continue
        h = r[3] - r[1]
        if 20 <= h <= 50:
            if top_y is None or r[1] < top_y:
                top_y = r[1]
                toolbar = ch

    if not toolbar:
        raise RuntimeError("Barra de ferramentas do visualizador não encontrada.")
    try: viewer.SetFocus()
    except: pass
    return viewer, toolbar

def _walk_buttons(node, acc):
    try:
        for c in node.GetChildren():
            if _tipo(c) == "ButtonControl":
                acc.append(c)
            _walk_buttons(c, acc)
    except:
        pass

def listar_botoes_toolbar(toolbar):
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
    try:
        inv = alvo.GetInvokePattern()
        if inv:
            inv.Invoke()
            print(f"[OK] Invoke no botão #{indice_from_1}.")
            return True
    except:
        pass
    try: alvo.SetFocus()
    except: pass
    uia.SendKeys("{Enter}")
    print(f"[OK] Foco+Enter no botão #{indice_from_1}.")
    return True

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

def wait_global_aviso_do_sistema(timeout=600, interval=0.25, max_depth=8):
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

# ===================== Relatório: pipeline (chamada única) =====================
def gerar_relatorio(pasta_xml: str):
    """
    Abre: Relatórios > Legais > Livro Registro de ISS > Serviços Prestados - Padrão
    Preenche MM/AAAA do mês anterior, clica Imprimir, na visualização clica o ícone (BOTAO_VIS_EXPORTAR_IDX),
    escreve 'pasta_xml', confirma o aviso global e finaliza.
    """
    root = uia.GetRootControl()

    # Abrir via menu
    workspace = bfs_find(root, WORKSPACE_NAME, types=('PaneControl','GroupControl','DocumentControl'), max_depth=4) or root
    rel = wait_until(lambda: find_first_by_subname([root, workspace], MENU_RELATORIOS,
                                                   types=('MenuItemControl','TabItemControl','ButtonControl'), max_depth=7),
                     timeout=20, interval=0.4)
    if not rel: raise RuntimeError("Menu 'Relatórios' não encontrado.")
    uia_activate(rel, "menu 'Relatórios'")

    legais = wait_until(lambda: find_first_by_subname([root, workspace], SUB_LEGAIS,
                                                      types=('MenuItemControl','ButtonControl','ListItemControl','TreeItemControl'), max_depth=8),
                        timeout=20, interval=0.4)
    if not legais: raise RuntimeError("Submenu 'Legais' não encontrado.")
    uia_activate(legais, "submenu 'Legais'")

    livro = wait_until(lambda: find_first_by_subname([root, workspace], ITEM_LIVRO_REG_ISS,
                                                     types=('MenuItemControl','ButtonControl','ListItemControl','TreeItemControl'), max_depth=9),
                       timeout=20, interval=0.4)
    if not livro: raise RuntimeError("Item 'Livro Registro de ISS' não encontrado.")
    uia_activate(livro, "item 'Livro Registro de ISS'")

    serv_padrao = wait_until(lambda: find_first_by_subname([root, workspace], ITEM_SERV_PREST_PADRAO,
                                                           types=('MenuItemControl','ButtonControl','ListItemControl','TreeItemControl'), max_depth=10),
                             timeout=20, interval=0.4)
    if not serv_padrao: raise RuntimeError("Item 'Serviços Prestados - Padrão' não encontrado.")
    uia_activate(serv_padrao, "item 'Serviços Prestados - Padrão'")

    # Espera a janela do relatório
    def localizar_janela_livro():
        fiscal = _find_fiscal_root()
        if not fiscal:
            return None
        workspace2 = None
        for ch in fiscal.GetChildren():
            if _tipo(ch) == "PaneControl" and _nome(ch) == WORKSPACE_NAME:
                workspace2 = ch; break
        if not workspace2:
            return None
        for ch in workspace2.GetChildren():
            if _tipo(ch) == "WindowControl" and _nome(ch) == JANELA_LIVRO:
                try: ch.SetFocus()
                except: pass
                return ch
        return None

    janela_livro = wait_until(localizar_janela_livro, timeout=25, interval=0.5)
    if not janela_livro:
        raise RuntimeError("Janela do Livro não localizada.")

    # Preenche e imprime
    try: janela_livro.SetFocus()
    except: pass
    time.sleep(0.2)
    pag.press('tab')
    pag.write(obter_mes_ano_anterior_slash())
    time.sleep(0.2)
    pag.press('tab')
    time.sleep(0.2)

    btn_imprimir = bfs_find(janela_livro, BTN_IMPRIMIR, types=('ButtonControl',), max_depth=8)
    if not btn_imprimir:
        raise RuntimeError("Botão 'Imprimir' não encontrado.")
    uia_activate(btn_imprimir, "botão 'Imprimir'")

    # Visualizador -> clicar ícone (ex.: exportar)
    time.sleep(0.6)
    _viewer, toolbar = localizar_visualizador_relatorio()
    listar_botoes_toolbar(toolbar)  # sempre lista (ajuda a calibrar índices)
    clicar_botao_por_indice(toolbar, BOTAO_VIS_EXPORTAR_IDX)

    # Caixa de diálogo para caminho/confirmar
    time.sleep(0.2)
    pag.write(pasta_xml)
    time.sleep(1.0)
    pag.press('tab'); pag.press('tab'); pag.press('tab'); pag.press('space')
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

    # (opcional) clicar outro botão na barra após salvar
    clicar_botao_por_indice(toolbar, BOTAO_VIS_APOS_SALVAR_IDX)

def wait_first_aviso_or_importar(janela_nfse, fiscal_win, timeout=6000, interval=0.25):
    """
    Espera simultaneamente pelo 'Aviso do Sistema' ou pelo botão 'Importar' habilitar.
    Retorna:
      ('aviso', ctrl_aviso, texto)  ou
      ('importar', btn_importar_habilitado, None)  ou
      None (timeout)
    """
    end = time.time() + timeout

    def importar_ready():
        _, b2 = encontrar_botoes_nfse(janela_nfse)
        b = b2
        if not b:
            return None
        bring_into_view(b)
        try:
            if hasattr(b, "IsEnabled") and not b.IsEnabled:
                return None
        except:
            pass
        try:
            if b.GetInvokePattern():
                return b
        except:
            return None
        return b

    while time.time() < end:
        # 1) Checa aviso (se já existir janela Fiscal)
        if fiscal_win:
            try:
                aviso, txt = wait_aviso_do_sistema(fiscal_win, timeout=0.01, interval=0.01)
            except:
                aviso, txt = (None, None)
            if aviso:
                return ('aviso', aviso, txt)

        # 2) Checa botão Importar habilitado
        btn_imp_ok = importar_ready()
        if btn_imp_ok:
            return ('importar', btn_imp_ok, None)

        time.sleep(interval)

    return None

# ===================== Fator / Processamento Especial =====================
def processar_empresa_com_fator(fatorR: str):
    """
    Função especial para processar empresas que possuem FATOR.
    Adicione aqui a lógica específica necessária.
    """
    print(f"\n[FATOR] Processando empresa COM FATOR...")
    
    """Abre menus sequencialmente com timeouts curtos."""
    root = uia.GetRootControl()
    workspace = bfs_find(root, WORKSPACE_NAME,
                        types=("PaneControl", "GroupControl", "DocumentControl"),
                        max_depth=4) or root

    trib = wait_until(
        lambda: find_first_by_subname([root, workspace], MENU_TRIBUTOS, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if trib:
        uia_activate_fast(trib)
        time.sleep(0.1)
    else:
        print("[AVISO] Menu 'Tributos' não encontrado")
        return

    simples = wait_until(
        lambda: find_first_by_subname([root, workspace], SUB_SIMPLES_NACIONAL, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if simples:
        uia_activate_fast(simples)
        time.sleep(0.1)
    else:
        print("[AVISO] Submenu 'Simples Nacional' não encontrado")
        return

    ITEM_FATOR_R = wait_until(
        lambda: find_first_by_subname([root, workspace], MENU_FATOR_R, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if ITEM_FATOR_R:
        uia_activate_fast(ITEM_FATOR_R)
        time.sleep(0.1)
    else:
        print("[AVISO] Item 'Fator R' não encontrado")
        return
    
    janela_fator = wait_until(localizar_fator_dentro_fiscal, timeout=25, interval=0.5)
    if not janela_fator:
        raise RuntimeError("Janela de Fator R não foi encontrada a tempo.")

    pag.write(obter_mes_ano_anterior_slash())
    pag.press('tab'); pag.press('f4')
    
    time.sleep(1)
    
    pag.press('alt+f'); pag.press('enter')
    
    btn_apuracao, btn_gravar = encontrar_botoes_fator(janela_fator)
    if not btn_apuracao:
        raise RuntimeError("Botão 'Apuração' não apareceu.")
    
    if not btn_gravar:
        raise RuntimeError("Botão 'Gravar' não encontrado.")
    
    # CARREGAR
    uia_activate(btn_apuracao, "botão 'Apuração'")
    time.sleep(0.2)
    
    pag.press('alt+g'); time.sleep(1); pag.press('space')
    
    time.sleep(3)
        
    uia_activate(btn_gravar, "botão 'Gravar'")

    print(f"[FATOR] ✓ Processamento com fator finalizado")

# ===================== PIPELINE PRINCIPAL =====================
def main():
    print(">>> (1) Lendo XMLs e montando LOG…")
    linhas_log = montar_log_empresas()
    print(">>> (2) Registrando LOG…")
    registrar_log(linhas_log)

    empresas = carregar_resultados_log(CAMINHO_LOG)
    if not empresas:
        print("Nenhuma empresa válida no LOG. Encerrando.")
        return

    print(">>> (3) Ativando e maximizando o ERP…")
    win_pyget = ativar_e_maximizar()
    if not win_pyget:
        print("[ERRO] Não consegui ativar a janela do Fiscal. Encerrando.")
        return
    time.sleep(0.8)

    win_uia_root = uia.GetRootControl()
    print(">>> (4) Iniciando importações…")

    for codigo_erp, estab, fatorR, pasta_xml in empresas:
        print(f"\n=== Empresa {codigo_erp} / Estab {estab} ===")

        print(">>> (5) Trocando empresa…")
        trocar_empresa(
            win_pyget,
            codigo=str(codigo_erp),
            estabelecimento=str(estab),
            data_ddmmaa=ultimo_dia_mes_anterior()
        )

        time.sleep(1)

        pag.press('space')


        # Abrir Importação → NFS-e
        workspace = bfs_find(win_uia_root, WORKSPACE_NAME,
                             types=('PaneControl','GroupControl','DocumentControl'),
                             max_depth=4) or win_uia_root

        menu_import = wait_until(
            lambda: find_first_by_subname([win_uia_root, workspace], "importação",
                                          types=('MenuItemControl','TabItemControl','ButtonControl'), max_depth=6),
            timeout=20, interval=0.5
        )
        if not menu_import:
            raise RuntimeError("Menu 'Importação' não encontrado.")
        uia_activate(menu_import, "menu 'Importação'")

        nfse_item = wait_until(
            lambda: find_first_by_subname([win_uia_root, workspace], "nfs-e (serviços prestados/tomados)",
                                          types=('MenuItemControl','ButtonControl','ListItemControl','TreeItemControl'),
                                          max_depth=7),
            timeout=20, interval=0.5
        )
        if not nfse_item:
            raise RuntimeError("Item 'NFS-e (Serviços Prestados/Tomados)' não encontrado.")
        uia_activate(nfse_item, "item NFS-e (Serviços Prestados/Tomados)")

        janela_nfse = wait_until(localizar_nfse_dentro_fiscal, timeout=60, interval=0.5)
        if not janela_nfse:
            #####
            pag.press('space')

            time.sleep(1)
            # Abrir Importação → NFS-e
            workspace = bfs_find(win_uia_root, WORKSPACE_NAME,
                                types=('PaneControl','GroupControl','DocumentControl'),
                                max_depth=4) or win_uia_root

            menu_import = wait_until(
                lambda: find_first_by_subname([win_uia_root, workspace], "importação",
                                            types=('MenuItemControl','TabItemControl','ButtonControl'), max_depth=6),
                timeout=20, interval=0.5
            )
            if not menu_import:
                raise RuntimeError("Menu 'Importação' não encontrado.")
            uia_activate(menu_import, "menu 'Importação'")

            nfse_item = wait_until(
                lambda: find_first_by_subname([win_uia_root, workspace], "nfs-e (serviços prestados/tomados)",
                                            types=('MenuItemControl','ButtonControl','ListItemControl','TreeItemControl'),
                                            max_depth=7),
                timeout=20, interval=0.5
            )
            if not nfse_item:
                raise RuntimeError("Item 'NFS-e (Serviços Prestados/Tomados)' não encontrado.")
            uia_activate(nfse_item, "item NFS-e (Serviços Prestados/Tomados)")

            janela_nfse = wait_until(localizar_nfse_dentro_fiscal, timeout=60, interval=0.5)
            if not janela_nfse:
                raise RuntimeError("Janela de NFS-e não foi encontrada a tempo.")
            ####
            raise RuntimeError("Janela de NFS-e não foi encontrada a tempo.")

        btn_carregar, btn_importar = encontrar_botoes_nfse(janela_nfse)
        if not btn_carregar:
            raise RuntimeError("Botão 'Carregar' não encontrado.")
        if not btn_importar:
            raise RuntimeError("Botão 'Importar' não apareceu.")

        # Preenche diretório e período
        pag.press('tab'); pag.write(pasta_xml); time.sleep(0.2)
        pag.press('tab'); pag.write(obter_mes_ano_anterior_slash()); time.sleep(0.2)

        # CARREGAR
        uia_activate(btn_carregar, "botão 'Carregar'")
        fiscal_win = get_fiscal_from_nfse(janela_nfse)

        # Espera concorrente: aviso OU importar habilitar (timeout=600)
        resultado = wait_first_aviso_or_importar(janela_nfse, fiscal_win, timeout=6000, interval=0.25)
        if not resultado:
            raise RuntimeError("Nem apareceu 'Aviso do Sistema' nem habilitou 'Importar' dentro do tempo limite (600s).")

        tipo, obj, texto = resultado

        if tipo == 'aviso':
            # Mesmo tratamento anterior do aviso depois de Carregar
            print("\n[AVISO DO SISTEMA - após CARREGAR]")
            print((texto or "").strip() if texto else "(sem texto)")
            focus_and_dismiss_alert(obj)
            pag.press('space')
            time.sleep(0.3)
            fechar_todas_as_janelas()
            continue  # NÃO gera relatório — apenas na importação

        # Se caiu aqui, o 'Importar' ficou habilitado primeiro
        btn_importar_enabled = obj

        # 3) IMPORTAR → pode surgir AVISO (incluindo erro da GDFe)
        sucesso_importar = tentar_importar_com_retry_gdfe(
            btn_importar_enabled,
            fiscal_win,
            max_tentativas=3  # ajuste se quiser mais/menos tentativas
        )

        if not sucesso_importar:
            print("[WARN] IMPORTAR não foi concluído (erro GDFe persistente ou falha em clicar). "
                  "Pulando geração de relatório para esta empresa.")
            try:
                fechar_todas_as_janelas()
            except Exception as e:
                print(f"[WARN] Não consegui clicar 'Sair' na NFS-e após falha no Importar: {e}")
            continue  # vai para a próxima empresa do loop

        # ====== GERAÇÃO DE RELATÓRIO (somente se Importar foi concluído) ======
        try:
            gerar_relatorio(pasta_xml)
        except Exception as e:
            print(f"[WARN] Falha ao gerar relatório: {e}")
            
        # Encerra tela NFS-e
        try:
            fechar_todas_as_janelas()
        except Exception as e:
            print(f"[WARN] Não consegui clicar 'Sair' na NFS-e: {e}")

                
        # ====== FATOR R ======
        sinal = str(fatorR or "").strip().lower()

        if sinal in ("sim"):
            try:
                processar_empresa_com_fator(fatorR)
            except Exception as e:
                print(f"[WARN] Falha ao processar FATOR R: {e}")
            finally:
                try:
                    fechar_todas_as_janelas()
                except Exception as e:
                    print(f"[WARN] Não consegui fechar janelas após FATOR R: {e}")
        else:
            print(f"[INFO] Fator R marcado como '{fatorR}' — pulando processamento de FATOR R.")



        print("\n>>> Fluxo finalizado para esta empresa.")

if __name__ == "__main__":
    main()
