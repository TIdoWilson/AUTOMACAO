# -*- coding: utf-8 -*-
"""
Fluxo:
1) Tareffa (Playwright): login -> filtro -> exportar CSV -> salvar em ./downloads
2) CSV: montar lista única por (empresa, estabelecimento) a partir das colunas "Característica" e "codigoERP"
3) ERP (UIAutomation): ativar "Folha de Pagamento" -> trocar empresa (apenas na primeira ou quando muda) 
   -> gerar "Resumo da folha mensal" e/ou "Relatório de Férias"

Observações:
- Datas: sempre mês/ano anterior (MM/AAAA) e último dia do mês anterior (DDMMAAAA).
- Processamento: 2
- Classificação: 1
- Marcar: "Listar recibos de pró-labore", "Listar recibos de autônomos", "Imprimir dados em destaque"
- Troca de empresa: apenas quando empresa muda; se só muda estabelecimento, apenas preenche campos
"""

from __future__ import annotations

import csv
import datetime as dt
import re
import time
import unicodedata
import os
import shutil
import ctypes
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Tuple
import traceback

import pyautogui as pag
import pygetwindow as gw
import uiautomation as uia
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ===================== TAREFFA (Playwright) =====================
LOGIN_URL = (
    "https://prd-ottimizza-oauth-server.ottimizza.dev/login"
    "?redirect=%252Foauth%252Fauthorize%252Foauthchooseaccount%253Fresponse_type%253Dcode"
    "%2526prompt%253Dlogin%2526client_id%253D9ae2b3e71b8ee9a0810b%2526redirect_uri%253D"
    "https%253A%252F%252Fweb.tareffa.com.br%252Fauth%252Fcallback"
)

EMPRESAS_URL = "https://web.tareffa.com.br/empresas"



def ler_credenciais(txt_path: Path) -> tuple[str, str]:
    txt = txt_path.read_text(encoding="utf-8", errors="ignore")
    m_email = re.search(r'TAREFFA_EMAIL\s*=\s*["\']([^"\']+)["\']', txt)
    m_senha = re.search(r'TAREFFA_SENHA\s*=\s*["\']([^"\']+)["\']', txt)
    if not m_email or not m_senha:
        raise ValueError('Esperado no txt: TAREFFA_EMAIL="..." e TAREFFA_SENHA="...".')
    return m_email.group(1).strip(), m_senha.group(1).strip()


def click_if_visible(page, locator, timeout_ms=1500) -> bool:
    try:
        locator.first.wait_for(state="visible", timeout=timeout_ms)
        locator.first.click()
        return True
    except PlaywrightTimeoutError:
        return False


def finalizar_oauth_choose_account(page):
    choose_btn = page.locator("body > div > div > div.form-body > ul > form > button").first
    choose_btn.wait_for(state="visible", timeout=5000)
    choose_btn.click()

    for label in ["Continuar", "Autorizar", "Permitir", "Prosseguir", "Acessar", "Confirmar", "OK"]:
        try:
            btn = page.get_by_role("button", name=re.compile(label, re.I))
            if btn.count() > 0:
                btn.first.click()
        except PlaywrightTimeoutError:
            pass

    page.wait_for_url(re.compile(r"https://web\.tareffa\.com\.br/.*"), timeout=5000)

    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PlaywrightTimeoutError:
        pass


# ===================== CSV -> Jobs únicos =====================
@dataclass(frozen=True)
class EmpresaJob:
    # Código ORIGINAL (como aparece no CSV / como deve sair no nome do PDF)
    empresa_original: int
    estabelecimento_original: int

    # Código para GERAÇÃO no ERP (pode ser convertido)
    empresa: int
    estabelecimento: int

    gerar_resumo_folha: bool
    gerar_relatorio_ferias: bool

    @property
    def codigo_erp_original(self) -> str:
        return f"{self.empresa_original}-{self.estabelecimento_original}"

    @property
    def codigo_erp_geracao(self) -> str:
        return f"{self.empresa}-{self.estabelecimento}"

def _read_sample_text(path: Path) -> str:
    b = path.read_bytes()[:12000]
    try:
        return b.decode("utf-8-sig")
    except UnicodeDecodeError:
        return b.decode("ISO-8859-1", errors="ignore")

def _detect_delimiter(path: Path) -> str:
    sample = _read_sample_text(path)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,|\t")
        return dialect.delimiter
    except Exception:
        return ";"

def _get_col(row: dict, *names: str) -> str:
    norm = {str(k).strip().lower(): k for k in row.keys()}
    for name in names:
        key = name.strip().lower()
        if key in norm:
            return str(row[norm[key]] or "").strip()
    return ""

def _parse_codigo_erp(codigo: str) -> tuple[int, int]:
    codigo = (codigo or "").strip()
    if not codigo:
        raise ValueError("codigoERP vazio no CSV.")

    if "-" in codigo:
        a, b = codigo.split("-", 1)
        empresa = int(re.sub(r"\D", "", a) or "0")
        estab = int(re.sub(r"\D", "", b) or "0")
        if empresa <= 0 or estab <= 0:
            raise ValueError(f"CódigoERP inválido: {codigo!r}")
        return empresa, estab

    empresa = int(re.sub(r"\D", "", codigo) or "0")
    if empresa <= 0:
        raise ValueError(f"CódigoERP inválido: {codigo!r}")
    return empresa, 1  # sem filial => estabelecimento 1


# ===================== Regras especiais de códigos =====================
# Empresas que devem gerar relatório usando estabelecimento igual ao número da empresa (X-Y -> X-X),
# mas manter o código original (do CSV) para nomear o PDF.
EMPRESAS_FORCAR_ESTAB_IGUAL_EMPRESA: set[int] = {39, 258, 91, 6, 246, 252, 4, 2, 166}

# Conversões pontuais (código original -> código para geração no ERP)
# Ex.: 187-3 deve gerar como 187-2, mas nome do PDF continua 187-3.
CONVERSOES_PONTUAIS: dict[tuple[int, int], tuple[int, int]] = {
    (187, 3): (187, 2), (660, 4): (660, 2),
}

# Lista de exclusão: empresas aqui NÃO serão geradas
EMPRESAS_EXCLUIR: set[int] = {121, 199, 324}

def montar_lista_processamento(csv_path: Path) -> list[EmpresaJob]:
    csv_path = Path(csv_path)
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV não encontrado: {csv_path}")

    delimiter = _detect_delimiter(csv_path)

    def iter_rows(enc: str):
        with csv_path.open("r", encoding=enc, newline="") as f:
            reader = csv.DictReader(f, delimiter=delimiter)
            for row in reader:
                yield row

    flags: dict[tuple[int, int], dict[str, bool]] = {}
    ordem: list[tuple[int, int]] = []
    conv_map: dict[tuple[int, int], tuple[int, int]] = {}

    def process_row(row: dict):
        codigo = _get_col(row, "codigoERP", "códigoerp", "CódigoERP")
        caract = _get_col(row, "característica", "Característica", "caracteristica", "Caracteristica")
        if not codigo or not caract:
            return

        c = caract.lower()
        tem_func = "tem funcionario" in c
        tem_pro = ("tem pro-labore" in c) or ("tem prolabore" in c) or ("tem pro labore" in c)
        if not (tem_func or tem_pro):
            return

        empresa_orig, estab_orig = _parse_codigo_erp(codigo)

        # Exclusão por empresa (não gera)
        if empresa_orig in EMPRESAS_EXCLUIR:
            return

        # Aplica conversões para geração no ERP
        empresa_gen, estab_gen = empresa_orig, estab_orig
        if (empresa_orig, estab_orig) in CONVERSOES_PONTUAIS:
            empresa_gen, estab_gen = CONVERSOES_PONTUAIS[(empresa_orig, estab_orig)]
        elif empresa_orig in EMPRESAS_FORCAR_ESTAB_IGUAL_EMPRESA:
            estab_gen = empresa_orig

        key = (empresa_orig, estab_orig)
        conv_map[key] = (empresa_gen, estab_gen)
        if key not in flags:
            flags[key] = {"tem_func": False, "tem_pro": False}
            ordem.append(key)
        flags[key]["tem_func"] = flags[key]["tem_func"] or tem_func
        flags[key]["tem_pro"] = flags[key]["tem_pro"] or tem_pro

    try:
        for row in iter_rows("utf-8-sig"):
            process_row(row)
    except UnicodeDecodeError:
        for row in iter_rows("ISO-8859-1"):
            process_row(row)

    jobs: list[EmpresaJob] = []
    for empresa, estab in ordem:
        tem_func = flags[(empresa, estab)]["tem_func"]
        tem_pro = flags[(empresa, estab)]["tem_pro"]
        jobs.append(
            EmpresaJob(
                empresa_original=empresa,
                estabelecimento_original=estab,
                empresa=conv_map[(empresa, estab)][0],
                estabelecimento=conv_map[(empresa, estab)][1],
                gerar_resumo_folha=(tem_func or tem_pro),
                gerar_relatorio_ferias=tem_func,
            )
        )
    return jobs
# ===================== Datas =====================
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

def obter_mes_ano_anterior_compacto(ref: dt.date | None = None) -> str:
    """Retorna data em formato MMAAAA (sem /)"""
    hoje = ref or dt.date.today()
    primeiro_do_mes = hoje.replace(day=1)
    ultimo = primeiro_do_mes - dt.timedelta(days=1)
    return f"{ultimo.month:02d}{ultimo.year:04d}"

def obter_mes_ano_compacto(ref: dt.date | None = None) -> str:
    """Retorna data em formato MMAAAA (sem /) do mês atual."""
    hoje = ref or dt.date.today()
    return f"{hoje.month:02d}{hoje.year:04d}"

def gerar_nome_arquivo_resumo_folha(job: EmpresaJob) -> str:
    """Gera nome do arquivo: 1-1 RESUMO DA FOLHA 012024.pdf (ou 1 RESUMO... quando estab=1)"""
    mes_ano = obter_mes_ano_anterior_compacto()
    empresa_fmt = str(job.empresa_original).lstrip('0') or '0'
    estab_fmt = str(job.estabelecimento_original).lstrip('0') or '0'

    prefixo = empresa_fmt if job.estabelecimento_original == 1 else f"{empresa_fmt}-{estab_fmt}"
    return f"{prefixo} RESUMO DA FOLHA {mes_ano}.pdf"

def gerar_nome_arquivo_relatorio_ferias(job: EmpresaJob) -> str:
    """Gera nome do arquivo: 1-1 RELATORIODEFERIAS 012024.pdf (ou 1 RELATORIO... quando estab=1)"""
    mes_ano = obter_mes_ano_compacto()
    empresa_fmt = str(job.empresa_original).lstrip('0') or '0'
    estab_fmt = str(job.estabelecimento_original).lstrip('0') or '0'

    prefixo = empresa_fmt if job.estabelecimento_original == 1 else f"{empresa_fmt}-{estab_fmt}"
    return f"{prefixo} RELATORIODEFERIAS {mes_ano}.pdf"

def preencher_pasta_e_nome_pdf(tipo: str, job: EmpresaJob):
    """
    Preenche o campo de pasta e o campo de nome do arquivo PDF no diálogo de salvar.
    tipo: "resumo" ou "ferias"
    """
    if tipo == "resumo":
        nome_arquivo = gerar_nome_arquivo_resumo_folha(job)
    elif tipo == "ferias":
        nome_arquivo = gerar_nome_arquivo_relatorio_ferias(job)
    else:
        raise ValueError(f"Tipo inválido: {tipo}")

    caminho_pasta = str(PASTA_SAIDA)
    
    print(f"[LOG] Preenchendo: Pasta='{caminho_pasta}' | Nome='{nome_arquivo}'")
    
    # Preenche o campo de caminho (primeira tab)
    pag.write(caminho_pasta, interval=0.01)
    time.sleep(0.2)
    pag.press('tab')
    time.sleep(0.2)
    
    # Preenche o campo de nome do arquivo (segunda tab)
    pag.write(nome_arquivo, interval=0.01)
    time.sleep(0.2)

# ===================== ERP (UIA) - Config =====================
FOLHA_TITLE_EXATO = "Folha de Pagamento"
GERENCIADOR_SISTEMAS_TITLE_SUBSTR = "Gerenciador de Sistemas"
WORKSPACE_NAME = "Espaço de trabalho"

MENU_RELATORIOS = "Relatórios"
SUB_MENSAIS = "Mensais"
ITEM_RESUMO_FOLHA = "Resumo da Folha mensal"

MENU_MODULOS = "Módulos"
SUB_FERIAS = "Férias"
ITEM_RELATORIO_DE_FERIAS = "Relatório de Férias"

MENU_JANELA = "Janela"
SUB_FECHAR_TODAS = "Fechar Todas"

BOTAO_IMPRIMIR = "Imprimir"
BOTAO_OK_IMPRIMIR = "OK"

# BoundingRectangle absolutos para Resumo da Folha
RECT_MESANO_DE = (770, 334, 816, 355)
RECT_MESANO_ATE = (829, 334, 875, 355)
RECT_PROCESSAMENTO = (770, 356, 816, 377)
RECT_CLASSIFICACAO = (770, 378, 816, 399)
RECT_ESTAB_DE = (770, 400, 816, 421)
RECT_ESTAB_ATE = (849, 400, 895, 421)
RECT_CHK_PROLABORE = (926, 418, 1077, 435)
RECT_CHK_AUTONOMOS = (926, 434, 1085, 451)
RECT_CHK_DESTAQUE = (926, 466, 1082, 483)

# BoundingRectangle absolutos para Relatório de Férias
RECT_CLASSIFICACAO1 = (786, 325, 819, 346)
RECT_ESTAB = (786, 371, 1126, 392)
RECT_CHK_DESTAQUE1 = (755, 702, 981, 719)

# PASTA PARA SALVAR RELATORIOS PDF
PASTA_SAIDA = Path(__file__).resolve().parent / "relatorios_pdf"
PASTA_SAIDA.mkdir(parents=True, exist_ok=True)


# Toolbar do visualizador — índices
BOTAO_VIS_EXPORTAR_IDX = 16
BOTAO_VIS_APOS_SALVAR_IDX = 13

pag.PAUSE = 0.05
pag.FAILSAFE = True

ROI_RIGHT_WIDTH = 560
ROI_TOP_OFFSET = 34
ROI_HEIGHT = 130
GRID_COLS = 10
GRID_ROWS = 6
MIN_W, MAX_W = 18, 160
MIN_H, MAX_H = 14, 36

# ===================== UIA Helpers (otimizados) =====================
def normalize(s: str) -> str:
    try:
        return s.lower().strip()
    except Exception:
        return ""

def wait_until(fn, timeout=10.0, interval=0.15):
    end = time.time() + timeout
    while time.time() < end:
        try:
            result = fn()
            if result:
                return result
        except Exception:
            pass
        time.sleep(interval)
    return None

def bfs_find(root_ctrl, name_substr: str,
             types: Tuple[str, ...] = (
                 "WindowControl", "PaneControl", "GroupControl", "DocumentControl",
                 "ButtonControl", "EditControl", "MenuItemControl", "ListItemControl",
                 "TreeItemControl", "TabItemControl"
             ),
             max_depth: int = 6):
    target = normalize(name_substr)
    q = deque([(root_ctrl, 0)])
    while q:
        node, depth = q.popleft()
        if depth > max_depth:
            continue
        try:
            if normalize(getattr(node, 'Name', '')) and target in normalize(getattr(node, 'Name', '')):
                if node.ControlTypeName in types:
                    return node
            for child in node.GetChildren():
                q.append((child, depth + 1))
        except Exception:
            pass
    return None

def find_first_by_subname(scopes, subname: str,
                          types: Tuple[str, ...],
                          max_depth: int = 6):
    for scope in scopes:
        result = bfs_find(scope, subname, types, max_depth)
        if result:
            return result
    return None

def rect_of(ctrl):
    try:
        r = ctrl.BoundingRectangle
        return (r.left, r.top, r.right, r.bottom) if r else None
    except Exception:
        return None

def bring_into_view(ctrl):
    try:
        sp = ctrl.GetScrollItemPattern()
        if sp:
            sp.ScrollIntoView()
    except Exception:
        pass
    try:
        wp = ctrl.GetWindowPattern()
        if wp:
            wp.SetWindowVisualState(0)
    except Exception:
        pass

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

def set_value_direct(ctrl, text: str) -> bool:
    """Escreve direto via ValuePattern."""
    if not ctrl:
        return False
    try:
        vp = ctrl.GetValuePattern()
        if vp:
            vp.SetValue(text)
            return True
    except Exception:
        pass
    return False

def toggle_checkbox(ctrl) -> bool:
    """Marca checkbox via TogglePattern."""
    if not ctrl:
        return False
    try:
        tp = ctrl.GetTogglePattern()
        if tp and tp.ToggleState != 1:
            tp.Toggle()
            return True
        elif tp:
            return True
    except Exception:
        pass
    return False

def center_of_rect(rect):
    l, t, r, b = rect
    return ((l + r) / 2.0, (t + b) / 2.0)

def build_control_index(root_ctrl, types: Tuple[str, ...], max_depth: int = 6):
    """Indexa controles uma vez (rápido)."""
    idx = []
    q = deque([(root_ctrl, 0)])
    while q:
        node, depth = q.popleft()
        if depth > max_depth:
            continue
        try:
            if node.ControlTypeName in types:
                r = rect_of(node)
                if r:
                    cx, cy = center_of_rect(r)
                    idx.append((node, r, (cx, cy)))
            for child in node.GetChildren():
                q.append((child, depth + 1))
        except Exception:
            pass
    return idx

def find_control_near_rect_fast(target_rect, control_type_name: str,
                                index, max_dist: float = 40.0):
    """Busca no índice (O(n), muito rápido)."""
    tx, ty = center_of_rect(target_rect)
    best = None
    best_dist = None
    for ctrl, rect, (cx, cy) in index:
        try:
            if ctrl.ControlTypeName != control_type_name:
                continue
        except Exception:
            continue
        dist = abs(cx - tx) + abs(cy - ty)
        if best is None or dist < best_dist:
            best = ctrl
            best_dist = dist
    
    if best and best_dist <= max_dist:
        return best
    return None

def _tipo(ctrl):
    try: return getattr(ctrl, "ControlTypeName", "")
    except: return ""
    
def _nome(ctrl):
    try: return (getattr(ctrl, "Name", "") or "").strip()
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
        if _tipo(ctrl) in ("TextControl","EditControl"):
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
        t = _tipo(cur)
        if t in ("DialogControl","WindowControl"):
            return cur
        cur = _parent(cur)
    return node

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
                ctype = _tipo(node)
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

# ===================== Ativar / maximizar Folha =====================
def _get_window_by_title_substr(substr: str):
    """Retorna a maior janela visível cujo título contém `substr` (case-insensitive)."""
    try:
        wins = [w for w in gw.getAllWindows() if w.visible and substr.lower() in (w.title or "").lower()]
    except Exception:
        wins = []
    if not wins:
        return None
    return max(wins, key=lambda x: x.width * x.height)


def reabrir_folha(empresa: int, estabelecimento: int):
    """
    Recuperação quando o ERP/folha some durante a execução.
    Passos (conforme solicitado):
      - Shift+S (fecha a folha)
      - Ativa 'Gerenciador de Sistemas'
      - Clica em 800x187
      - Aguarda 10s
      - Alt+N
      - Reposiciona na empresa/estabelecimento atual
    """
    print("[RECUP] Tentando reabrir a Folha de Pagamento...")

    # 1) Fecha/limpa a folha (atalho)
    try:
        pag.hotkey('shift', 's')
    except Exception as e:
        print(f"[RECUP][WARN] Falha ao enviar Shift+S: {e}")

    time.sleep(0.8)

    # 2) Ativa o Gerenciador de Sistemas
    gerente = _get_window_by_title_substr(GERENCIADOR_SISTEMAS_TITLE_SUBSTR) or _get_window_by_title_substr("gerenciador")
    if not gerente:
        print("[RECUP][ERRO] Janela 'Gerenciador de Sistemas' não encontrada.")
    else:
        try:
            if not gerente.isActive:
                gerente.activate()
                time.sleep(0.3)
        except Exception as e:
            print(f"[RECUP][WARN] Erro ao ativar Gerenciador via pygetwindow: {e}")

    # 3) Clique coordenado e aguarda
    try:
        pag.moveTo(800, 187, duration=0.15)
        pag.click(800, 187)
    except Exception as e:
        print(f"[RECUP][WARN] Falha ao clicar em 800x187: {e}")

    time.sleep(10)

    # 4) Alt+N
    try:
        pag.hotkey('alt', 'n')
    except Exception as e:
        print(f"[RECUP][WARN] Falha ao enviar Alt+N: {e}")

    # 5) Volta para a Folha e reconfigura empresa atual
    time.sleep(3.0)
    win_pyget = ativar_e_maximizar()
    if not win_pyget:
        raise RuntimeError("[RECUP] Não consegui reativar a janela 'Folha de Pagamento' após reabrir.")

    try:
        trocar_empresa(
            win_pyget,
            codigo=str(empresa),
            estabelecimento=str(estabelecimento),
            data_ddmmaa=ultimo_dia_mes_anterior(),
        )
        print(f"[RECUP] ✓ Empresa reconfigurada: {empresa}-{estabelecimento}")
    except Exception as e:
        raise RuntimeError(f"[RECUP] Falha ao reconfigurar empresa após reabrir: {e}") from e

    return win_pyget

def ativar_e_maximizar():
    wins = [w for w in gw.getWindowsWithTitle(FOLHA_TITLE_EXATO) if w.visible]
    if not wins:
        print("[ERRO] Nenhuma janela 'Folha de Pagamento' encontrada.")
        return None
    w = max(wins, key=lambda x: x.width * x.height)
    print(f"[LOG] Ativando: {w.title!r}")
    
    # Tenta ativar com tratamento de erro
    try:
        if not w.isActive:
            w.activate()
            time.sleep(0.3)
    except Exception as e:
        print(f"[WARN] Erro ao ativar via pygetwindow: {e}")
        print(f"[LOG] Tentando alternativa com ctypes...")
        try:
            import ctypes
            hwnd = ctypes.windll.user32.FindWindowW(None, w.title)
            if hwnd:
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                time.sleep(0.3)
        except Exception as e2:
            print(f"[WARN] Erro na alternativa: {e2}")
    
    # Tenta maximizar
    try:
        import ctypes
        from ctypes import wintypes
        SPI_GETWORKAREA = 48
        rect = wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
        w.moveTo(rect.left, rect.top)
        w.resizeTo(rect.right - rect.left, rect.bottom - rect.top)
    except Exception as e:
        print(f"[WARN] Erro ao maximizar: {e}")
        try:
            w.maximize()
        except Exception as e2:
            print(f"[WARN] Falha na maximização: {e2}")
    
    time.sleep(0.1)
    return w

def _find_Folha_root() -> Optional[uia.Control]:
    root = uia.GetRootControl()
    for w in root.GetChildren():
        if FOLHA_TITLE_EXATO.lower() in (getattr(w, 'Name', '') or '').lower():
            return w
    return None
    
# ===================== Abrir menus =====================
def abrir_tela_resumo_Folha():
    """Abre menus sequencialmente com timeouts curtos."""
    root = uia.GetRootControl()
    workspace = bfs_find(root, WORKSPACE_NAME,
                        types=("PaneControl", "GroupControl", "DocumentControl"),
                        max_depth=4) or root

    rel = wait_until(
        lambda: find_first_by_subname([root, workspace], MENU_RELATORIOS, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if rel:
        uia_activate_fast(rel)
        time.sleep(0.1)
    else:
        print("[AVISO] Menu 'Relatórios' não encontrado")
        return

    mensais = wait_until(
        lambda: find_first_by_subname([root, workspace], SUB_MENSAIS, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if mensais:
        uia_activate_fast(mensais)
        time.sleep(0.1)
    else:
        print("[AVISO] Submenu 'Mensais' não encontrado")
        return

    item_resumo = wait_until(
        lambda: find_first_by_subname([root, workspace], ITEM_RESUMO_FOLHA, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if item_resumo:
        uia_activate_fast(item_resumo)
        time.sleep(0.1)
    else:
        print("[AVISO] Item 'Resumo da Folha mensal' não encontrado")

def abrir_tela_relatorio_de_ferias():
    """Abre menus sequencialmente com timeouts curtos."""
    root = uia.GetRootControl()
    workspace = bfs_find(root, WORKSPACE_NAME,
                        types=("PaneControl", "GroupControl", "DocumentControl"),
                        max_depth=4) or root

    rel = wait_until(
        lambda: find_first_by_subname([root, workspace], MENU_MODULOS, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if rel:
        uia_activate_fast(rel)
        time.sleep(0.1)
    else:
        print("[AVISO] Menu 'Módulos' não encontrado")
        return

    ferias = wait_until(
        lambda: find_first_by_subname([root, workspace], SUB_FERIAS, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if ferias:
        uia_activate_fast(ferias)
        time.sleep(0.1)
    else:
        print("[AVISO] Submenu 'Férias' não encontrado")
        return

    item_ferias = wait_until(
        lambda: find_first_by_subname([root, workspace], ITEM_RELATORIO_DE_FERIAS, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if item_ferias:
        uia_activate_fast(item_ferias)
        time.sleep(0.1)
    else:
        print("[AVISO] Item 'Relatório de Férias' não encontrado")

def fechar_todas_as_janelas():
    """Fecha todas as janelas abertas no ERP."""
    # Pré-checagem: se já existir um "Aviso do Sistema" pendente, fecha e confirma 2x (Alt+S)
    dlg_pre, txt_pre = wait_global_aviso_do_sistema(timeout=5, interval=0.25, max_depth=8)
    if dlg_pre:
        print(f"[LOG] Aviso pendente encontrado antes de 'Fechar Todas': {txt_pre if txt_pre else '(sem texto)'}")
        try:
            dlg_pre.SetFocus()
        except Exception:
            pass
        time.sleep(0.2)
        pag.hotkey('alt', 's')
        time.sleep(0.2)
        pag.hotkey('alt', 's')
        time.sleep(0.2)

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
        time.sleep(0.1)
    else:
        print("[AVISO] Menu 'Janela' não encontrado")
        return

    janela = wait_until(
        lambda: find_first_by_subname([root, workspace], SUB_FECHAR_TODAS, ("MenuItemControl", "ButtonControl"), 6),
        timeout=5, interval=0.1
    )
    if janela:
        uia_activate_fast(janela)
        time.sleep(0.3)  # Aguarda o diálogo aparecer
    else:
        print("[AVISO] Submenu 'Fechar Todas' não encontrado")
        return
    
    # Aguarda o aviso do sistema
    print("[LOG] Aguardando confirmação do aviso...")
    dlg, dlg_text = wait_global_aviso_do_sistema(timeout=10, interval=0.25, max_depth=8)
    if dlg:
        print(f"[AVISO DO SISTEMA] {dlg_text if dlg_text else '(sem texto)'}")
        try:
            dlg.SetFocus()
        except:
            pass
        time.sleep(0.2)
        # Pressiona Alt+S para confirmar (Sim)
        pag.hotkey('alt', 's')
        print("[LOG] ✓ Diálogo confirmado")
        time.sleep(0.5)
    else:
        print("[WARN] Aviso do Sistema não apareceu ou timeout expirou")

def trocar_empresa(win_window, codigo: str, estabelecimento: str, data_ddmmaa: str):
    if not win_window:
        raise RuntimeError("Janela do ERP não está ativa.")

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
                        r = rect_of(el)
                        if r and r not in seen:
                            seen.add(r)
                            candidatos.append(el)
                        break
                    el = el.GetParentControl()
            except Exception:
                continue

    if not candidatos:
        raise RuntimeError("Não encontrei campos na ROI. Ajuste ROI_* se necessário.")

    filtrados = [c for c in candidatos if is_edit_candidate(c)]
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
        mesma_linha = [c for c in filtrados if abs(rect_of(c)[1] - rect_of(filtrados[0])[1]) <= 10]
        campo_data = sorted(mesma_linha, key=lambda c: rect_of(c)[0])[-1]

    if len(restantes) < 2:
        restantes = [c for c in filtrados if c is not campo_data]
    restantes.sort(key=lambda c: (rect_of(c)[1], rect_of(c)[0]))
    campo_codigo, campo_estab = restantes[0], (restantes[1] if len(restantes) > 1 else None)

    set_text(campo_codigo, codigo)
    try:
        campo_codigo.SetFocus()
    except Exception:
        pass
    uia.SendKeys("{Enter}")
    uia.SendKeys("{Enter}")
    time.sleep(1.0)

    if campo_estab:
        set_text(campo_estab, estabelecimento)

    set_text(campo_data, data_ddmmaa)
    uia.SendKeys("{Enter}")
    uia.SendKeys("{Enter}")
    time.sleep(0.8)

# ===================== Toolbar / PDF =====================
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

def localizar_visualizador_relatorio(timeout: float = 300.0, interval: float = 5.0):
    """
    Tenta localizar o visualizador do relatório a cada `interval` segundos,
    até `timeout` segundos. Retorna (viewer, toolbar) ou lança RuntimeError.
    """
    def _attempt():
        try:
            Folha = _find_Folha_root()
            if not Folha:
                return None
            workspace = None
            for ch in Folha.GetChildren():
                if _tipo(ch) == "PaneControl" and _nome(ch) == WORKSPACE_NAME:
                    workspace = ch
                    break
            if not workspace:
                return None
            viewer = None
            for ch in workspace.GetChildren():
                if _tipo(ch) == "WindowControl" and (_nome(ch) == "" or "Visualizador" in _nome(ch)):
                    viewer = ch
                    break
            if not viewer:
                return None

            toolbar, top_y = None, None
            for ch in viewer.GetChildren():
                if _tipo(ch) != "PaneControl":
                    continue
                r = rect_of(ch)
                if not r:
                    continue
                h = r[3] - r[1]
                if 20 <= h <= 60:
                    if top_y is None or r[1] < top_y:
                        top_y = r[1]
                        toolbar = ch

            if not toolbar:
                return None
            try:
                viewer.SetFocus()
            except:
                pass
            return viewer, toolbar
        except Exception:
            return None

    resultado = wait_until(_attempt, timeout=timeout, interval=interval)
    if resultado:
        return resultado
    raise RuntimeError(f"Visualizador do relatório não encontrado após {timeout} segundos.")

def fechar_aviso_do_sistema(dlg):
    """Fecha um aviso do sistema pressionando Alt+S (Sim/OK)."""
    try:
        if dlg:
            dlg.SetFocus()
    except:
        pass
    time.sleep(0.2)
    pag.hotkey('space')
    time.sleep(0.5)

def aguardar_visualizador_ou_aviso(timeout: float = 300.0, interval: float = 2.0):
    """
    Loop que aguarda o visualizador do relatório ou um aviso do sistema.
    Retorna:
    - ("visualizador", viewer, toolbar) se encontrar o visualizador
    - ("aviso", dlg, txt) se encontrar um aviso
    - ("timeout", None, None) se expirar o timeout
    """
    end = time.time() + timeout
    
    while time.time() < end:
        # Tenta localizar o visualizador
        try:
            Folha = _find_Folha_root()
            if Folha:
                workspace = None
                for ch in Folha.GetChildren():
                    if _tipo(ch) == "PaneControl" and _nome(ch) == WORKSPACE_NAME:
                        workspace = ch
                        break
                if workspace:
                    viewer = None
                    for ch in workspace.GetChildren():
                        if _tipo(ch) == "WindowControl" and (_nome(ch) == "" or "Visualizador" in _nome(ch)):
                            viewer = ch
                            break
                    if viewer:
                        toolbar, top_y = None, None
                        for ch in viewer.GetChildren():
                            if _tipo(ch) != "PaneControl":
                                continue
                            r = rect_of(ch)
                            if not r:
                                continue
                            h = r[3] - r[1]
                            if 20 <= h <= 60:
                                if top_y is None or r[1] < top_y:
                                    top_y = r[1]
                                    toolbar = ch
                        if toolbar:
                            try:
                                viewer.SetFocus()
                            except:
                                pass
                            print("[LOG] ✓ Visualizador do relatório localizado!")
                            return ("visualizador", viewer, toolbar)
        except Exception as e:
            print(f"[LOG] Erro ao localizar visualizador: {e}")
        
        # Tenta localizar aviso do sistema
        try:
            root = uia.GetRootControl()
            q = deque([(root, 0)])
            aviso_encontrado = False
            
            while q and not aviso_encontrado:
                node, d = q.popleft()
                if d > 8:
                    continue
                try:
                    ctype = _tipo(node)
                    nm = _nome(node)
                    if ctype in ("DialogControl", "WindowControl") and nm == "Aviso do Sistema":
                        txt = "\n".join(_collect_texts(node)) or nm
                        print(f"[AVISO] Aviso do Sistema encontrado: {txt}")
                        return ("aviso", node, txt)
                    if ctype == "TitleBarControl" and _value(node) == "Aviso do Sistema":
                        dlg = _dialog_ancestor(node)
                        txt = "\n".join(_collect_texts(dlg)) or "Aviso do Sistema"
                        print(f"[AVISO] Aviso do Sistema encontrado: {txt}")
                        return ("aviso", dlg, txt)
                except:
                    pass
                try:
                    for ch in node.GetChildren():
                        q.append((ch, d + 1))
                except:
                    pass
        except Exception as e:
            print(f"[LOG] Erro ao localizar aviso: {e}")
        
        time.sleep(interval)
    
    print(f"[TIMEOUT] Visualizador/Aviso não encontrado após {timeout}s")
    return ("timeout", None, None)

# ===================== Gerar Resumo Folha (otimizado) =====================
def gerar_resumo_folha_mensal(job: EmpresaJob):
    """Gera resumo da folha mensal para um job específico."""
    print(f"\n[LOG] === Gerando Resumo da Folha para {job.codigo_erp_original} ===")
    
    main_win = _find_Folha_root()
    if not main_win:
        raise RuntimeError("Janela UIA 'Folha de Pagamento' não localizada.")

    print(f"[LOG] Janela UIA: {getattr(main_win, 'Name', '')!r}")
    try:
        main_win.SetFocus()
    except Exception:
        pass
    time.sleep(0.1)

    print("[LOG] Abrindo menu 'Resumo da Folha mensal'...")
    abrir_tela_resumo_Folha()

    janela = _find_Folha_root() or main_win
    print(f"[LOG] Container: {getattr(janela, 'Name', '')!r}")
    try:
        janela.SetFocus()
    except Exception:
        pass
    time.sleep(0.1)

    mes_ano = obter_mes_ano_anterior_slash()
    print(f"[LOG] Competência: {mes_ano}")

    print("[LOG] Indexando controles...")
    edits_index = build_control_index(janela, ("EditControl",), max_depth=6)
    checks_index = build_control_index(janela, ("CheckBoxControl",), max_depth=6)

    print(f"[LOG] Indexados: {len(edits_index)} EditControls, {len(checks_index)} CheckBoxes")
    time.sleep(0.1)

    controls = {}
    for key, rect, ctrl_type, index in [
        ("mes_de", RECT_MESANO_DE, "EditControl", edits_index),
        ("mes_ate", RECT_MESANO_ATE, "EditControl", edits_index),
        ("estab_de", RECT_ESTAB_DE, "EditControl", edits_index),
        ("estab_ate", RECT_ESTAB_ATE, "EditControl", edits_index),
        ("proc", RECT_PROCESSAMENTO, "EditControl", edits_index),
        ("class", RECT_CLASSIFICACAO, "EditControl", edits_index),
        ("chk_pro", RECT_CHK_PROLABORE, "CheckBoxControl", checks_index),
        ("chk_aut", RECT_CHK_AUTONOMOS, "CheckBoxControl", checks_index),
        ("chk_dest", RECT_CHK_DESTAQUE, "CheckBoxControl", checks_index),
    ]:
        controls[key] = find_control_near_rect_fast(rect, ctrl_type, index, max_dist=25.0)
        if controls[key]:
            print(f"  ✓ {key}")
        else:
            print(f"  ✗ {key}")

    print("[LOG] Preenchendo campos...")
    if controls.get("mes_de"):
        set_value_direct(controls["mes_de"], mes_ano)
    if controls.get("mes_ate"):
        set_value_direct(controls["mes_ate"], mes_ano)
    if controls.get("proc"):
        set_value_direct(controls["proc"], "2")
    if controls.get("class"):
        set_value_direct(controls["class"], "1")
    if controls.get("estab_de"):
        set_value_direct(controls["estab_de"], str(job.estabelecimento))
    if controls.get("estab_ate"):
        set_value_direct(controls["estab_ate"], str(job.estabelecimento))


    print("[LOG] Marcando checkboxes...")
    toggle_checkbox(controls.get("chk_pro"))
    toggle_checkbox(controls.get("chk_aut"))
    toggle_checkbox(controls.get("chk_dest"))

    time.sleep(0.5)

    print("[LOG] Localizando botão 'Imprimir'...")
    root = uia.GetRootControl()
    botao_imprimir = wait_until(
        lambda: find_first_by_subname([root, janela], BOTAO_IMPRIMIR, ("ButtonControl",), 6),
        timeout=5, interval=0.1
    )

    if botao_imprimir:
        print("[LOG] Botão 'Imprimir' encontrado. Acionando...")
        if uia_activate_fast(botao_imprimir, "Imprimir"):
            print("[LOG] ✓ Botão 'Imprimir' acionado com sucesso!")
            time.sleep(2.0)

    # Loop aguardando visualizador ou aviso
    print("[LOG] Aguardando visualizador ou aviso do sistema...")
    tipo, obj1, obj2 = aguardar_visualizador_ou_aviso(timeout=300.0, interval=2.0)
    
    if tipo == "timeout":
        print("[ERRO] Timeout: visualizador não aparecer e nenhum aviso detectado!")
        raise RuntimeError("Visualizador do relatório não localizado.")
    
    if tipo == "aviso":
        dlg, txt = obj1, obj2
        print(f"[LOG] Aviso detectado, fechando: {txt}")
        fechar_aviso_do_sistema(dlg)
        print("[LOG] ✓ Pulando para próximo relatório...")
        return  # Retorna sem gerar o PDF
    
    # tipo == "visualizador"
    viewer, toolbar = obj1, obj2
    
    listar_botoes_toolbar(toolbar)
    clicar_botao_por_indice(toolbar, BOTAO_VIS_EXPORTAR_IDX)
    
    time.sleep(0.2)
    # Preenche pasta e nome do arquivo PDF
    preencher_pasta_e_nome_pdf("resumo", job)

    pag.press('tab'); pag.press('tab'); pag.press('space')
    pag.press('tab'); pag.press('space'); pag.press('tab'); pag.press('space')

    dlg, dlg_text = wait_global_aviso_do_sistema(timeout=30, interval=0.25, max_depth=8)
    if dlg:
        print("\n[AVISO DO SISTEMA - relatório]")
        print(dlg_text if dlg_text else "(sem texto)")
        try: dlg.SetFocus()
        except: pass

    pag.hotkey('alt','s')
    time.sleep(0.5)
    pag.press('space')

    clicar_botao_por_indice(toolbar, BOTAO_VIS_APOS_SALVAR_IDX)
    caminho_pdf = str(PASTA_SAIDA / gerar_nome_arquivo_resumo_folha(job))
    print(f"[LOG] ✓ Resumo da Folha gerado e salvo em: {caminho_pdf}")

# ===================== Gerar Relatório de Férias (otimizado) =====================
def gerar_relatorio_ferias(job: EmpresaJob):
    """Gera relatório de férias para um job específico."""
    print(f"\n[LOG] === Gerando Relatório de Férias para {job.codigo_erp_original} ===")
    
    main_win = _find_Folha_root()
    if not main_win:
        raise RuntimeError("Janela UIA 'Folha de Pagamento' não localizada.")

    print(f"[LOG] Janela UIA: {getattr(main_win, 'Name', '')!r}")
    try:
        main_win.SetFocus()
    except Exception:
        pass
    time.sleep(0.1)

    print("[LOG] Abrindo menu 'Relatório de Férias'...")
    abrir_tela_relatorio_de_ferias()

    janela = _find_Folha_root() or main_win
    print(f"[LOG] Container: {getattr(janela, 'Name', '')!r}")
    try:
        janela.SetFocus()
    except Exception:
        pass
    time.sleep(0.1)

    print("[LOG] Indexando controles...")
    edits_index = build_control_index(janela, ("EditControl",), max_depth=6)
    checks_index = build_control_index(janela, ("CheckBoxControl",), max_depth=6)

    print(f"[LOG] Indexados: {len(edits_index)} EditControls, {len(checks_index)} CheckBoxes")
    time.sleep(0.1)

    controls = {}
    for key, rect, ctrl_type, index in [
        ("class", RECT_CLASSIFICACAO1, "EditControl", edits_index),
        ("estab", RECT_ESTAB, "EditControl", edits_index),
        ("chk_dest", RECT_CHK_DESTAQUE1, "CheckBoxControl", checks_index),
    ]:
        controls[key] = find_control_near_rect_fast(rect, ctrl_type, index, max_dist=25.0)
        if controls[key]:
            print(f"  ✓ {key}")
        else:
            print(f"  ✗ {key}")

    print("[LOG] Preenchendo campos...")
    if controls.get("class"):
        set_value_direct(controls["class"], "1")
    if controls.get("estab"):
        set_value_direct(controls["estab"], str(job.estabelecimento))

    print("[LOG] Marcando checkboxes...")
    toggle_checkbox(controls.get("chk_dest"))

    time.sleep(0.5)

    print("[LOG] Localizando botão 'Imprimir'...")
    root = uia.GetRootControl()
    botao_imprimir = wait_until(
        lambda: find_first_by_subname([root, janela], BOTAO_IMPRIMIR, ("ButtonControl",), 6),
        timeout=5, interval=0.1
    )

    if botao_imprimir:
        print("[LOG] Botão 'Imprimir' encontrado. Acionando...")
        if uia_activate_fast(botao_imprimir, "Imprimir"):
            print("[LOG] ✓ Botão 'Imprimir' acionado com sucesso!")
            time.sleep(2.0)

    # Loop aguardando visualizador ou aviso
    print("[LOG] Aguardando visualizador ou aviso do sistema...")
    tipo, obj1, obj2 = aguardar_visualizador_ou_aviso(timeout=300.0, interval=2.0)
    
    if tipo == "timeout":
        print("[ERRO] Timeout: visualizador não aparecer e nenhum aviso detectado!")
        raise RuntimeError("Visualizador do relatório não localizado.")
    
    if tipo == "aviso":
        dlg, txt = obj1, obj2
        print(f"[LOG] Aviso detectado, fechando: {txt}")
        fechar_aviso_do_sistema(dlg)
        print("[LOG] ✓ Pulando para próximo relatório...")
        return  # Retorna sem gerar o PDF
    
    # tipo == "visualizador"
    viewer, toolbar = obj1, obj2
    
    listar_botoes_toolbar(toolbar)
    clicar_botao_por_indice(toolbar, BOTAO_VIS_EXPORTAR_IDX)

    time.sleep(0.2)
    # Preenche pasta e nome do arquivo PDF
    preencher_pasta_e_nome_pdf("ferias", job)

    pag.press('tab'); pag.press('tab'); pag.press('space')
    pag.press('tab'); pag.press('space'); pag.press('tab'); pag.press('space')

    dlg, dlg_text = wait_global_aviso_do_sistema(timeout=30, interval=0.25, max_depth=8)
    if dlg:
        print("\n[AVISO DO SISTEMA - relatório]")
        print(dlg_text if dlg_text else "(sem texto)")
        try: dlg.SetFocus()
        except: pass

    pag.hotkey('alt','s')
    time.sleep(0.2)
    pag.press('space')

    clicar_botao_por_indice(toolbar, BOTAO_VIS_APOS_SALVAR_IDX)
    caminho_pdf = str(PASTA_SAIDA / gerar_nome_arquivo_relatorio_ferias(job))
    print(f"[LOG] ✓ Relatório de Férias gerado e salvo em: {caminho_pdf}")

# ===================== Executar jobs no ERP =====================
def executar_jobs_no_erp(jobs: Iterable[EmpresaJob]):
    """
    Executa jobs agrupados por empresa.
    - Primeira vez que vê uma empresa: abre janela de troca
    - Mesma empresa, diferente estabelecimento: apenas preenche estab_de/estab_ate/estab
    - Fecha todas as janelas ANTES de cada geração de relatório
    """
    win_pyget = ativar_e_maximizar()
    if not win_pyget:
        win_pyget = reabrir_folha(jobs[0].empresa, jobs[0].estabelecimento)

    empresa_atual = None

    for job in jobs:
        print(f"\n{'='*60}")
        print(f"[LOG] Job: {job.codigo_erp_original}")
        print(f"      Resumo Folha: {job.gerar_resumo_folha} | Férias: {job.gerar_relatorio_ferias}")
        print(f"{'='*60}")

        # Sempre fecha as janelas abertas ANTES de gerar relatório
        print(f"[LOG] Fechando janelas abertas...")
        try:
            fechar_todas_as_janelas()
        except Exception as e:
            print(f"[WARN] Erro ao fechar janelas: {e}")
        time.sleep(0.5)

        # Verifica se precisa trocar de empresa
        if job.empresa != empresa_atual:
            print(f"[LOG] 🔄 Empresa mudou: {empresa_atual} → {job.empresa}")
            print(f"[LOG] Abrindo tela de troca de empresa...")

            trocar_empresa(
                win_pyget,
                codigo=str(job.empresa),
                estabelecimento=str(job.estabelecimento),
                data_ddmmaa=ultimo_dia_mes_anterior(),
            )

            print(f"[LOG] ✓ Empresa alterada para {job.empresa}")
            empresa_atual = job.empresa
            time.sleep(1.0)
        else:
            print(f"[LOG] 🔕 Mesma empresa ({job.empresa}), apenas alterando estabelecimento")


        # Gera resumo da folha se necessário
        if job.gerar_resumo_folha:
            try:
                gerar_resumo_folha_mensal(job)
            except Exception as e:
                print(f"[ERRO] Ao gerar Resumo Folha: {e}")
                traceback.print_exc()
            time.sleep(1.0)

        # Gera relatório de férias se necessário
        if job.gerar_relatorio_ferias:
            try:
                gerar_relatorio_ferias(job)
            except Exception as e:
                print(f"[ERRO] Ao gerar Relatório Férias: {e}")
                traceback.print_exc()
            time.sleep(1.0)

    print(f"\n{'='*60}")
    print(f"[LOG] ✓ Todos os jobs foram processados!")
    print(f"{'='*60}")

def baixar_csv_empresas_tareffa(email: str, senha: str, downloads_dir: Path) -> Path:
    """Baixa o CSV 'Empresas & Características' no Tareffa e retorna o caminho do arquivo salvo."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.set_default_timeout(60000)

        page.goto(LOGIN_URL, wait_until="domcontentloaded")
        page.get_by_role("textbox", name="Email").fill(email)
        page.get_by_role("textbox", name="Senha").fill(senha)
        page.get_by_role("button", name="Entrar").click()

        try:
            page.wait_for_url(re.compile(r".*(oauthchooseaccount|web\.tareffa\.com\.br).*"), timeout=5000)
        except PlaywrightTimeoutError:
            pass

        if "oauthchooseaccount" in page.url:
            finalizar_oauth_choose_account(page)
            page.wait_for_url(re.compile(r"https://web.tareffa.com.br/.*"), timeout=60000)

        page.goto(EMPRESAS_URL, wait_until="load")

        page.locator("button:has(mat-icon:has-text('more_vert'))").first.click()
        page.get_by_role("menuitem", name="Limpar Filtro").click()

        page.locator("button:has(mat-icon:has-text('more_vert'))").first.click()
        page.get_by_role("menuitem", name="Filtro Avançado").click()

        combo = page.get_by_role("combobox", name="Características")
        combo.fill("Tem pro-labore")
        page.get_by_role("option", name="Tem pro-labore").click()
        combo.press("Enter")

        combo.fill("Tem funcionario")
        page.get_by_role("option", name="Tem funcionario").click()
        combo.press("Enter")

        page.get_by_role("button", name="Aplicar").click()
        page.locator("#onlyActive").get_by_text("Somente em atividade").click()

        page.locator("button:has(mat-icon:has-text('more_vert'))").first.click()
        page.get_by_role("menuitem", name="Exportar para CSV").click()

        page.locator("#mat-select-value-7").click()
        page.get_by_role("option", name="Empresas & Características").click()

        with page.expect_download() as dlinfo:
            page.get_by_role("button", name="Download").click()

        download = dlinfo.value
        suggested = download.suggested_filename or "export.csv"
        csv_out = downloads_dir / suggested
        download.save_as(csv_out)
        _ = download.path()

        print(f"[OK] CSV baixado: {csv_out}")

        context.close()
        browser.close()

        return csv_out

def escolher_csv_base(email: str, senha: str, downloads_dir: Path) -> Path:
    """
    Pergunta ao iniciar se deve atualizar a base (baixar CSV novo no Tareffa).
    - Se SIM: baixa e salva/atualiza o cache fixo em ./downloads/empresas_caracteristicas.csv
    - Se NÃO: usa o cache salvo anteriormente.
    """
    downloads_dir.mkdir(parents=True, exist_ok=True)
    cache_file = downloads_dir / "empresas_caracteristicas.csv"

    resp = input("[INPUT] Atualizar a base de empresas (baixar novo CSV do Tareffa)? [S/N] (Enter=N): ").strip().lower()
    atualizar = resp in ("s", "sim", "y", "yes", "1", "true")

    if atualizar:
        csv_baixado = baixar_csv_empresas_tareffa(email, senha, downloads_dir)
        # Atualiza o cache de forma robusta (copy -> replace)
        try:
            tmp = cache_file.with_suffix(cache_file.suffix + ".tmp")
            shutil.copy2(csv_baixado, tmp)
            os.replace(tmp, cache_file)
            print(f"[OK] Base atualizada e salva em: {cache_file}")
            return cache_file
        except Exception as e:
            try:
                if tmp.exists():
                    tmp.unlink(missing_ok=True)
            except Exception:
                pass
            print(f"[WARN] Falha ao salvar cache ({e}); usando CSV baixado: {csv_baixado}")
            return csv_baixado

    # Não atualizar: usa o que já foi salvo anteriormente
    if cache_file.exists():
        print(f"[OK] Usando base salva anteriormente: {cache_file}")
        return cache_file

    # Se o cache não existir, tenta usar o CSV mais recente em ./downloads (melhor esforço)
    csvs = sorted(downloads_dir.glob("*.csv"), key=lambda f: f.stat().st_mtime, reverse=True)
    if csvs:
        print(f"[OK] Cache não encontrado; usando CSV mais recente em ./downloads: {csvs[0]}")
        return csvs[0]

    # Nada salvo: baixa um novo para não travar o fluxo
    print("[WARN] Nenhum CSV salvo encontrado; vou baixar um novo.")
    csv_baixado = baixar_csv_empresas_tareffa(email, senha, downloads_dir)
    try:
        tmp = cache_file.with_suffix(cache_file.suffix + ".tmp")
        shutil.copy2(csv_baixado, tmp)
        os.replace(tmp, cache_file)
        print(f"[OK] Base salva em: {cache_file}")
        return cache_file
    except Exception:
        return csv_baixado

# ===================== Main =====================
def main():
    base_dir = Path(__file__).resolve().parent
    email, senha = ler_credenciais(base_dir / "Senha Tareffa.txt")

    downloads_dir = base_dir / "downloads"
    downloads_dir.mkdir(parents=True, exist_ok=True)

    csv_out = escolher_csv_base(email, senha, downloads_dir)
    jobs = montar_lista_processamento(csv_out)
    print(f"[OK] Itens únicos para processar no ERP: {len(jobs)}")
    for i, job in enumerate(jobs, 1):
        print(f"  {i}. {job.codigo_erp_original} -> {job.codigo_erp_geracao} | Resumo: {job.gerar_resumo_folha} | Férias: {job.gerar_relatorio_ferias}")
    
    if not jobs:
        return

    # ===== NOVO: Pedir empresa/estabelecimento de início =====
    print(f"\n{'='*60}")
    inicio_user = input("[INPUT] Digite 'empresa-estabelecimento' para iniciar a partir dessa (ex: 1-1)\n"
                       "        ou deixe em branco/0 para começar do início: ").strip()
    print(f"{'='*60}\n")

    jobs_filtrados = jobs
    if inicio_user and inicio_user != "0":
        try:
            # Parse: "empresa-estab"
            partes = inicio_user.split("-")
            if len(partes) != 2:
                raise ValueError("Formato inválido. Use: empresa-estabelecimento")
            
            empresa_inicio = int(partes[0].strip())
            estab_inicio = int(partes[1].strip())
            
            # Encontra o índice onde começa
            idx_inicio = None
            for i, job in enumerate(jobs):
                if job.empresa == empresa_inicio and job.estabelecimento == estab_inicio:
                    idx_inicio = i
                    break
            
            if idx_inicio is None:
                print(f"[AVISO] Empresa-Estabelecimento '{inicio_user}' não encontrado na lista.")
                print(f"[LOG] Começando do início mesmo assim...")
            else:
                jobs_filtrados = jobs[idx_inicio:]
                print(f"[LOG] ✓ Iniciando a partir de {inicio_user}")
                print(f"[LOG] Total de jobs a processar: {len(jobs_filtrados)} (de {len(jobs)})\n")
        
        except ValueError as e:
            print(f"[ERRO] {e}")
            print(f"[LOG] Começando do início...")

    executar_jobs_no_erp(jobs_filtrados)

if __name__ == "__main__":
    main()
