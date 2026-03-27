import re
import glob
import time
import unicodedata
import pandas as pd
from pathlib import Path
from collections import OrderedDict, Counter
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE_EMPRESA_URL = "https://web.tareffa.com.br/empresas"

# =========================
# Utils
# =========================

def norm_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def strip_accents_upper(s: str) -> str:
    s = norm_spaces(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.upper()

def only_digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def to_id_str(v) -> str | None:
    """Converte o valor da planilha (pode vir como float/str) para ID em string (somente dígitos)."""
    if v is None:
        return None
    s = str(v).strip()
    if not s or s.lower() in {"nan", "none"}:
        return None
    # Excel pode trazer 35234132.0
    m = re.match(r"^(\d+)(?:\.0+)?$", s)
    if m:
        return m.group(1)
    # remove tudo que não for dígito
    digits = only_digits(s)
    return digits or None

def wait_dom(page):
    page.wait_for_load_state("domcontentloaded")

# =========================
# Login (mesma lógica dos seus códigos)
# =========================

def fill_email(page, value):
    candidatos = [
        page.get_by_label(re.compile(r"e-?mail", re.I)),
        page.get_by_placeholder(re.compile(r"e-?mail", re.I)),
        page.locator('input[type="email"]'),
        page.locator('input[name="email"]'),
        page.locator('input[formcontrolname="email"]'),
        page.locator('input[autocomplete="username"]'),
    ]
    for loc in candidatos:
        if loc.count():
            try:
                loc.first.wait_for(state="visible")
                loc.first.fill(value)
                return
            except Exception:
                continue
    raise RuntimeError("Campo de e-mail não encontrado.")

def fill_password(page, value):
    candidatos = [
        page.get_by_label(re.compile(r"senha|password", re.I)),
        page.get_by_placeholder(re.compile(r"senha|password", re.I)),
        page.locator('input[type="password"]'),
        page.locator('input[name="password"]'),
        page.locator('input[formcontrolname="password"]'),
        page.locator('input[autocomplete="current-password"]'),
    ]
    for loc in candidatos:
        if loc.count():
            try:
                loc.first.wait_for(state="visible")
                loc.first.fill(value)
                return
            except Exception:
                continue
    raise RuntimeError("Campo de senha não encontrado.")

def click_entrar(page):
    candidatos = [
        page.get_by_role("button", name=re.compile(r"^\s*entrar\s*$", re.I)),
        page.locator("button").filter(has_text=re.compile(r"\bentrar\b", re.I)),
        page.locator('button[type="submit"]'),
    ]
    for loc in candidatos:
        if loc.count():
            try:
                loc.first.wait_for(state="visible")
                loc.first.click()
                return
            except Exception:
                continue
    raise RuntimeError("Botão 'ENTRAR' não encontrado.")

def click_overlay_submit(page):
    alvo = page.locator('button[type="submit"][style*="position: absolute"][style*="z-index: 10"]')
    if alvo.count():
        try:
            alvo.first.wait_for(state="visible")
            alvo.first.click()
            return
        except Exception:
            pass
    comum = page.locator('button[type="submit"]')
    if comum.count():
        try:
            comum.first.wait_for(state="visible")
            comum.first.click()
            return
        except Exception:
            pass
    # não é crítico se não existir

# =========================
# Leitura da LISTA DOMESTICAS — AGRUPAR POR ID
# (A=Nome, B=CPF/CNPJ, C=ERP, ID=coluna 'ID' explícita)
# =========================

def carregar_domesticas_agrupado_por_id():
    base = Path(__file__).resolve().parent
    candidatos = []
    for ext in ("xlsx", "xls", "csv"):
        candidatos.extend(glob.glob(str(base / f"DOMESTICAS tareffa*.{ext}")))
    if not candidatos:
        raise FileNotFoundError("Arquivo 'DOMESTICAS tareffa.*' não encontrado.")

    caminho = sorted(candidatos)[0]
    suf = Path(caminho).suffix.lower()
    if suf in (".xlsx", ".xls"):
        df = pd.read_excel(caminho, dtype=str, header=0)
    elif suf == ".csv":
        try:
            df = pd.read_csv(caminho, dtype=str, header=0)
        except Exception:
            df = pd.read_csv(caminho, dtype=str, header=0, sep=";")
    else:
        raise ValueError(f"Extensão não suportada: {suf}")

    df = df.fillna("")

    # Detecta coluna ID (mantém sua lógica atual)
    id_col = None
    for c in df.columns:
        if str(c).strip().lower() == "id":
            id_col = c
            break
    if id_col is None:
        candidatos_id = []
        for c in df.columns:
            vals = df[c].astype(str).tolist()
            digits_ratio = sum(1 for v in vals if only_digits(v)) / max(1, len(vals))
            if digits_ratio > 0.7:
                candidatos_id.append(c)
        if candidatos_id:
            id_col = candidatos_id[-1]
        elif df.shape[1] >= 5:
            id_col = df.columns[-1]
        else:
            raise ValueError("Coluna 'ID' não encontrada na planilha.")

    # Esperado: A=Nome, B=CPF/CNPJ, C=ERP
    if df.shape[1] < 4:
        raise ValueError("Planilha precisa ter pelo menos 4 colunas (A=Nome, B=CPF/CNPJ, D=erp).")

    nomes = df.iloc[:, 0]
    cpfs  = df.iloc[:, 1]
    erp  = df.iloc[:, 3]
    ids   = df[id_col]

    # ↳ PRESERVA ORDEM DA PLANILHA
    grupos = OrderedDict()  # id -> {"id":..., "nomes":Counter, "cpf":digits, "erps":[...], "_seen":set()}

    for nome, cpf, erp, vid in zip(nomes, cpfs, erp, ids):
        nome = norm_spaces(str(nome))
        cpf  = norm_spaces(str(cpf))
        erp  = norm_spaces(str(erp))
        vid  = to_id_str(vid)
        if not vid or not nome or not cpf or not erp:
            continue

        if vid not in grupos:
            grupos[vid] = {
                "id": vid,
                "nomes": Counter(),
                "cpf": only_digits(cpf),
                "erps": [],        # mantém ordem de aparição
                "_seen": set(),    # evita duplicados de erp
            }

        g = grupos[vid]
        g["nomes"][nome] += 1
        if erp not in g["_seen"]:
            g["_seen"].add(erp)
            g["erps"].append(erp)

    if not grupos:
        raise ValueError("Nenhuma linha válida encontrada (ID, Nome, CPF/CNPJ e erp obrigatórios).")

    # Saída na MESMA ORDEM da planilha (sem sort!)
    out = []
    for g in grupos.values():
        nome_escolhido = g["nomes"].most_common(1)[0][0] if g["nomes"] else ""
        out.append({
            "id": g["id"],
            "nome": nome_escolhido,
            "cpf": g["cpf"],
            "erps": g["erps"],
        })
    return out

# =========================
# Página da Empresa – aba Características e adição (versões robustas)
# =========================

# 1) Espera o shell/tabs da página da empresa
def wait_empresa_tabs_ready(page, timeout_ms=30000):
    """
    Aguarda a estrutura de abas (header e corpos) da página de empresa.
    Usa apenas seletores genéricos do Angular Material.
    """
    page.wait_for_url(re.compile(r"https://web\.tareffa\.com\.br/empresas/\d+"), timeout=timeout_ms)
    page.wait_for_selector('.mat-mdc-tab-labels .mdc-tab__text-label', timeout=timeout_ms)
    page.wait_for_selector('.mat-mdc-tab-body', timeout=timeout_ms)

# 3) Abrir Características (um clique) + fallback de alternância se vier vazio
def ir_para_aba_caracteristicas(page, timeout_ms=25000, retries=2, toggle_wait_ms=350):
    """
    Ativa SEMPRE a aba 'Características' e espera SOMENTE o input de 'Adicionar serviço'.
    Se vier vazia, alterna para outra aba e volta (até 'retries' vezes).
    """
    serv_rx = re.compile(r"Características", re.I)

    # garante que as abas existem
    try:
        page.wait_for_selector('.mat-mdc-tab-labels .mdc-tab__text-label', timeout=30000)
        page.wait_for_selector('.mat-mdc-tab-body', timeout=30000)
    except PWTimeout:
        pass

    def _tabela_disponivel() -> bool:
        cont = page.locator(".mat-mdc-tab-body.mat-mdc-tab-body-active").first
        if not cont.count():
            return False
        # agora verificamos a presença da tabela de características
        tabela_loc = cont.locator(
            "mat-table.mat-mdc-table.mdc-data-table__table.cdk-table[role='table']"
        )
        return tabela_loc.count() > 0

    def _ativar_servicos():
        tab = page.get_by_role("tab", name=serv_rx)
        if not tab.count():
            tab = page.locator(
                '.mat-mdc-tab-labels .mdc-tab .mdc-tab__text-label',
                has_text=serv_rx
            )
        t = tab.first
        try:
            t.scroll_into_view_if_needed()
        except Exception:
            pass
        t.click()
        # tentar confirmar seleção e dar um respiro de rede
        try:
            page.wait_for_function(
                "el => el && el.getAttribute('aria-selected') === 'true'",
                arg=t, timeout=2000
            )
        except PWTimeout:
            pass
        try:
            page.wait_for_load_state("networkidle", timeout=1000)
        except PWTimeout:
            pass

    def _ativar_primeira_aba_diferente_de_servicos():
        tabs = page.get_by_role("tab")
        if tabs.count():
            for i in range(tabs.count()):
                lbl = tabs.nth(i)
                try:
                    txt = (lbl.inner_text() or "").strip()
                except Exception:
                    continue
                if not serv_rx.search(txt):
                    try:
                        lbl.scroll_into_view_if_needed()
                    except Exception:
                        pass
                    lbl.click()
                    return True
        # fallback por label cru
        labels = page.locator('.mat-mdc-tab-labels .mdc-tab .mdc-tab__text-label')
        for i in range(labels.count()):
            node = labels.nth(i)
            try:
                txt = (node.inner_text() or "").strip()
            except Exception:
                continue
            if not serv_rx.search(txt):
                try:
                    node.scroll_into_view_if_needed()
                except Exception:
                    pass
                node.click()
                return True
        return False

    # 1) SEMPRE ativa a aba Características
    _ativar_servicos()
        
    # 2) Espera o INPUT de Características
    inicio = time.time()
    while (time.time() - inicio) * 1000 < timeout_ms:
        if _tabela_disponivel():
            return
        time.sleep(0.2)

    # 3) Fallback: alterna outra aba e volta
    for _ in range(retries):
        if not _ativar_primeira_aba_diferente_de_servicos():
            break
        time.sleep(toggle_wait_ms / 1000.0)
        _ativar_servicos()

        inicio2 = time.time()
        while (time.time() - inicio2) * 1000 < timeout_ms // 2:
            if _tabela_disponivel():
                return
            time.sleep(0.2)

    raise TimeoutError("A aba 'Características' não exibiu o INPUT de 'Adicionar serviço' mesmo após alternar de aba e voltar.")

# 4) Pega o input (garantindo a aba)

# =========================
# Fluxo principal — navegação direta por ID
# =========================

# 6) Navegação encapsulada: ir direto ao ID e abrir Características
def abrir_empresa_e_caract(page, emp_id):
    page.goto(f"https://web.tareffa.com.br/empresas/{emp_id}")
    wait_dom(page)
    # Garante tabs do shell prontos ANTES de tentar mudar de aba
    wait_empresa_tabs_ready(page, timeout_ms=45000)
    # Agora sim, ativa Características (com fallback automático)
    ir_para_aba_caracteristicas(page, timeout_ms=25000, retries=2)


def selecionar_departamento_pessoal(page):
    # 1. Espera a tabela carregar e clica na linha desejada
    tabela = page.locator(
        "mat-table.mat-mdc-table.mdc-data-table__table.cdk-table[role='table']"
    )
    tabela.wait_for()  # espera a tabela renderizar
    linha = tabela.locator("mat-row:has-text('06. Departamento Pessoal')")
    linha.first.click()

    # 2. Aguarda a caixa de seleção aparecer
    def marcar_opcao_tem_domestica():
        checkbox = page.locator("mat-checkbox:has(span:has-text('Tem domestica'))")
        checkbox.wait_for()

        # Busca o <input type="checkbox"> interno
        input_cb = checkbox.locator("input[type='checkbox']")
        
        if input_cb.is_checked():
            print("✅ checkbox já marcada")

        # Só clica se estiver desmarcado
        if not input_cb.is_checked():
            input_cb.click()
            
    marcar_opcao_tem_domestica()

def fluxo():
    grupos = carregar_domesticas_agrupado_por_id()

    with sync_playwright() as p:
        navegador = p.chromium.launch(channel="chrome", headless=False)
        ctx = navegador.new_context(viewport={"width": 1920, "height": 1080})
        page = ctx.new_page()
        page.set_default_timeout(45000)

        # Login (uma vez)
        page.goto(BASE_EMPRESA_URL)
        wait_dom(page)
        fill_email(page, "contabil32@wilsonlopes.com.br")
        fill_password(page, "123456")
        click_entrar(page)
        wait_dom(page)
        try:
            click_overlay_submit(page)
        except Exception:
            pass
        
        # Espera determinística pela troca de URL para /servicos_programados
        page.wait_for_url(
            re.compile(r"^https://web\.tareffa\.com\.br/servicos_programados(?:/)?(?:\?.*)?(?:#.*)?$"),
            timeout=45000
        )
        wait_dom(page)  # garante domcontentloaded

        for i, grupo in enumerate(grupos, start=1):
            nome = grupo["nome"]
            cpf  = grupo["cpf"]
            erps = grupo["erps"]
            emp_id = grupo["id"]
            print(f"[{i}/{len(grupos)}] Empresa ID={emp_id} | {nome} | CPF/CNPJ: {cpf} | {len(erps)} serviço(s)")

            abrir_empresa_e_caract(page, emp_id)
            time.sleep(0.5)
            selecionar_departamento_pessoal(page)
            wait_dom(page)
            
        print("✅ Vinculação de Características por ID concluída.")

if __name__ == "__main__":
    fluxo()
