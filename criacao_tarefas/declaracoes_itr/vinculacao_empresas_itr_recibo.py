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
# Leitura da LISTA ITRS — AGRUPAR POR ID
# (A=Nome, B=CPF/CNPJ, D=CIB, ID=coluna 'ID' explícita)
# =========================

from collections import OrderedDict, Counter

def carregar_itrs_agrupado_por_id():
    base = Path(__file__).resolve().parent
    candidatos = []
    for ext in ("xlsx", "xls", "csv"):
        candidatos.extend(glob.glob(str(base / f"LISTA ITRS*.{ext}")))
    if not candidatos:
        raise FileNotFoundError("Arquivo 'LISTA ITRS.*' não encontrado.")

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

    # Esperado: A=Nome, B=CPF/CNPJ, D=CIB
    if df.shape[1] < 4:
        raise ValueError("Planilha precisa ter pelo menos 4 colunas (A=Nome, B=CPF/CNPJ, D=CIB).")

    nomes = df.iloc[:, 0]
    cpfs  = df.iloc[:, 1]
    cibs  = df.iloc[:, 3]
    ids   = df[id_col]

    # ↳ PRESERVA ORDEM DA PLANILHA
    grupos = OrderedDict()  # id -> {"id":..., "nomes":Counter, "cpf":digits, "cibs":[...], "_seen":set()}

    for nome, cpf, cib, vid in zip(nomes, cpfs, cibs, ids):
        nome = norm_spaces(str(nome))
        cpf  = norm_spaces(str(cpf))
        cib  = norm_spaces(str(cib))
        vid  = to_id_str(vid)
        if not vid or not nome or not cpf or not cib:
            continue

        if vid not in grupos:
            grupos[vid] = {
                "id": vid,
                "nomes": Counter(),
                "cpf": only_digits(cpf),
                "cibs": [],        # mantém ordem de aparição
                "_seen": set(),    # evita duplicados de CIB
            }

        g = grupos[vid]
        g["nomes"][nome] += 1
        if cib not in g["_seen"]:
            g["_seen"].add(cib)
            g["cibs"].append(cib)

    if not grupos:
        raise ValueError("Nenhuma linha válida encontrada (ID, Nome, CPF/CNPJ e CIB obrigatórios).")

    # Saída na MESMA ORDEM da planilha (sem sort!)
    out = []
    for g in grupos.values():
        nome_escolhido = g["nomes"].most_common(1)[0][0] if g["nomes"] else ""
        out.append({
            "id": g["id"],
            "nome": nome_escolhido,
            "cpf": g["cpf"],
            "cibs": g["cibs"],
        })
    return out

# =========================
# Página da Empresa – aba Serviços e adição (versões robustas)
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

# 2) Helper: container da aba ativa
def _servicos_container(page):
    cont = page.locator(".mat-mdc-tab-body.mat-mdc-tab-body-active").first
    cont.wait_for(state="attached", timeout=15000)
    return cont


# 3) Abrir ServiÇos (um clique) + fallback de alternância se vier vazio
def ir_para_aba_servicos(page, timeout_ms=25000, retries=2, toggle_wait_ms=350):
    """
    Ativa SEMPRE a aba 'Serviços' e espera SOMENTE o input de 'Adicionar serviço'.
    Se vier vazia, alterna para outra aba e volta (até 'retries' vezes).
    """
    serv_rx = re.compile(r"Servi[cç]os", re.I)

    # garante que as abas existem
    try:
        page.wait_for_selector('.mat-mdc-tab-labels .mdc-tab__text-label', timeout=30000)
        page.wait_for_selector('.mat-mdc-tab-body', timeout=30000)
    except PWTimeout:
        pass

    def _input_disponivel() -> bool:
        cont = page.locator(".mat-mdc-tab-body.mat-mdc-tab-body-active").first
        if not cont.count():
            return False
        # seletores mais específicos ao input de serviços
        input_loc = cont.locator(
            '#autoCompleteId, '
            'mat-form-field:has(mat-label:has-text("Adicionar serviço")) input, '
            'mat-form-field:has(mat-label:has-text("Adicionar serviço a empresa")) input'
        )
        return input_loc.count() > 0

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

    # 1) SEMPRE ativa a aba Serviços
    _ativar_servicos()

    # 2) Espera o INPUT de serviços
    inicio = time.time()
    while (time.time() - inicio) * 1000 < timeout_ms:
        if _input_disponivel():
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
            if _input_disponivel():
                return
            time.sleep(0.2)

    raise TimeoutError("A aba 'Serviços' não exibiu o INPUT de 'Adicionar serviço' mesmo após alternar de aba e voltar.")

# 4) Pega o input (garantindo a aba)
def _pegar_input_servico(page):
    # garante que estamos com controles presentes
    cont = _servicos_container(page)
    loc = cont.locator(
        '#autoCompleteId, '
        'input.mat-mdc-autocomplete-trigger[role="combobox"], '
        'mat-form-field:has(mat-label:has-text("Adicionar serviço")) input'
    )
    if not loc.count():
        raise TimeoutError("Input do serviço não está disponível na aba 'Serviços'.")
    loc.first.wait_for(state="visible", timeout=8000)
    return loc.first

def preencher_descricao_servico(page, texto):
    for _ in range(5):
        try:
            caixa = _pegar_input_servico(page)
            try:
                caixa.scroll_into_view_if_needed()
            except Exception:
                pass
            caixa.click()
            try:
                caixa.press("Control+A"); caixa.press("Delete")
            except Exception:
                pass
            try:
                caixa.fill("")
            except Exception:
                pass
            caixa.type(texto, delay=25)
            try:
                listbox = page.get_by_role("listbox")
                if listbox.count():
                    listbox.first.wait_for(state="visible", timeout=4000)
                    caixa.press("Enter")
            except PWTimeout:
                pass
            _pegar_input_servico(page).wait_for(state="visible", timeout=3000)
            return
        except Exception as e:
            if "not attached" in str(e).lower() or "não está disponível" in str(e).lower():
                time.sleep(0.25)
                continue
            raise
    raise RuntimeError("Não consegui estabilizar o input de serviço na aba 'Serviços'.")


# 5) Click Adicionar com revalidação de aba ativa
def click_adicionar_servico(page):
    # revalida que a aba está correta e com controles
    cont = _servicos_container(page)
    btn = cont.get_by_role("button", name=re.compile(r"^\s*Adicionar\s*$", re.I))
    if not btn.count():
        btn = cont.locator('button[mat-stroked-button]').filter(has_text=re.compile(r"\bAdicionar\b", re.I))

    # Se ainda assim não há botão, tenta reabrir Serviços (carregamento preguiçoso)
    if not btn.count():
        ir_para_aba_servicos(page, retries=2)
        cont = _servicos_container(page)
        btn = cont.get_by_role("button", name=re.compile(r"^\s*Adicionar\s*$", re.I))
        if not btn.count():
            btn = cont.locator('button[mat-stroked-button]').filter(has_text=re.compile(r"\bAdicionar\b", re.I))

    if not btn.count():
        raise TimeoutError("Botão 'Adicionar' não encontrado na aba 'Serviços'.")

    try:
        btn.first.scroll_into_view_if_needed()
    except Exception:
        pass
    btn.first.click()

    # Espera estabilizar: input costuma ser limpo após adicionar
    try:
        page.wait_for_function(
            """() => {
                const cont = document.querySelector('.mat-mdc-tab-body.mat-mdc-tab-body-active');
                if (!cont) return false;
                const el = cont.querySelector('#autoCompleteId, input.mat-mdc-autocomplete-trigger[role="combobox"]');
                return !!el && (!el.value || el.value.length === 0);
            }""",
            timeout=6000
        )
    except PWTimeout:
        time.sleep(0.3)

# =========================
# Fluxo principal — navegação direta por ID
# =========================

# 6) Navegação encapsulada: ir direto ao ID e abrir Serviços
def abrir_empresa_e_servicos(page, emp_id):
    page.goto(f"https://web.tareffa.com.br/empresas/{emp_id}")
    wait_dom(page)
    # Garante tabs do shell prontos ANTES de tentar mudar de aba
    wait_empresa_tabs_ready(page, timeout_ms=45000)
    # Agora sim, ativa Serviços (com fallback automático)
    ir_para_aba_servicos(page, timeout_ms=25000, retries=2)
    
def fluxo():
    grupos = carregar_itrs_agrupado_por_id()

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
            cibs = grupo["cibs"]
            emp_id = grupo["id"]
            print(f"[{i}/{len(grupos)}] Empresa ID={emp_id} | {nome} | CPF/CNPJ: {cpf} | {len(cibs)} serviço(s)")

            abrir_empresa_e_servicos(page, emp_id)
            
            time.sleep(0.5)

            # 3) Adiciona todos os serviços "RECIBO ITR {nome} {cib}"
            for cib in cibs:
                descricao = f"{nome} {cib}"
                preencher_descricao_servico(page, descricao)
                time.sleep(0.6)
                click_adicionar_servico(page)
                time.sleep(0.4)

            # Sem voltar para /empresas; já seguimos para o próximo ID

        print("✅ Vinculação de serviços por ID concluída.")


if __name__ == "__main__":
    fluxo()
