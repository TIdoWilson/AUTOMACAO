import re
import time
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# =====================================
# Configurações de URL (Empresas)
# =====================================
URL = "https://web.tareffa.com.br/empresas"
NEW_URL = "https://web.tareffa.com.br/empresas/nova"

# =====================================
# Helpers de espera
# =====================================

def wait_dom(page):
    page.wait_for_load_state("domcontentloaded")

def wait_new_empresa_form_ready(page):
    """
    Garante que o formulário de 'Nova Empresa' foi renderizado.
    Considera campos típicos: Código ERP, Nome Fantasia, Razão Social.
    """
    seletor = (
        'input[placeholder*="Código ERP" i], '
        'input[placeholder*="Nome Fantasia" i], '
        'input[placeholder*="Razão Social" i], '
        'mat-form-field:has(mat-label:has-text("Código ERP")) input, '
        'mat-form-field:has(mat-label:has-text("Nome Fantasia")) input, '
        'mat-form-field:has(mat-label:has-text("Razão Social")) input'
    )
    page.wait_for_selector(seletor, timeout=45000)

def open_nova_empresa(page):
    """Abre diretamente o formulário de 'Nova Empresa' e espera o form renderizar."""
    page.goto(NEW_URL)
    wait_dom(page)
    wait_new_empresa_form_ready(page)

# =====================================
# Preenchimento de campos
# =====================================
# =========================
# Login
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
    # Botão absoluto com z-index (se existir)
    alvo = page.locator(
        'button[type="submit"][style*="position: absolute"][style*="z-index: 10"]'
    )
    if alvo.count():
        try:
            alvo.first.wait_for(state="visible")
            alvo.first.click()
            return
        except Exception:
            pass
    # Fallback: qualquer submit
    comum = page.locator('button[type="submit"]')
    if comum.count():
        try:
            comum.first.wait_for(state="visible")
            comum.first.click()
            return
        except Exception:
            pass
    raise RuntimeError("Botão submit (overlay) não encontrado.")

def _count_safe(locator):
    try:
        return locator.count()
    except Exception:
        return 0

def _fill_first(page, locator, value):
    if _count_safe(locator):
        el = locator.first
        try:
            el.scroll_into_view_if_needed()
        except Exception:
            pass
        el.wait_for(state="visible", timeout=8000)
        el.fill(value)
        return True
    return False

def fill_codigo_erp(page, erp):
    candidatos = [
        page.get_by_placeholder(re.compile(r"^\s*C[oó]digo\s*ERP\s*$", re.I)),
        page.locator('mat-form-field:has(mat-label:has-text("Código ERP")) input'),
    ]
    for c in candidatos:
        if _fill_first(page, c, erp):
            return
    raise RuntimeError("Campo 'Código ERP' não encontrado.")

def fill_nome_fantasia(page, nome):
    candidatos = [
        page.get_by_placeholder(re.compile(r"^\s*Nome\s*Fantasia\s*$", re.I)),
        page.locator('mat-form-field:has(mat-label:has-text("Nome Fantasia")) input'),
    ]
    for c in candidatos:
        if _fill_first(page, c, nome):
            return
    raise RuntimeError("Campo 'Nome Fantasia' não encontrado.")

def fill_razao_social(page, nome):
    valor = f"{nome} - (FALTA USUÁRIOS)"
    candidatos = [
        page.get_by_placeholder(re.compile(r"Raz[aã]o\s*Social", re.I)),
        page.locator('mat-form-field:has(mat-label:has-text("Razão Social")) input'),
    ]
    for c in candidatos:
        if _fill_first(page, c, valor):
            return
    raise RuntimeError("Campo 'Razão Social' não encontrado.")

def fill_cpf_cnpj(page, cpf):
    # Mantém apenas dígitos; o componente normalmente mascara
    digits = re.sub(r"\D", "", str(cpf)).strip()

    candidatos = [
        # ✅ regex correta via engine semântico
        page.get_by_placeholder(re.compile(r"CPF\s*/\s*CNPJ\s*/\s*CNO\s*/\s*CAEPF", re.I)),
        # ✅ igualdade exata de placeholder (sem regex)
        page.locator('[placeholder="CPF/CNPJ/CNO/CAEPF"]'),
        # ✅ por label do mat-form-field (cobre variações)
        page.get_by_label(re.compile(r"CPF|CNPJ|CNO|CAEPF", re.I)),
        page.locator('mat-form-field:has(mat-label:has-text("CPF")) input'),
        page.locator('mat-form-field:has(mat-label:has-text("CPF/CNPJ")) input'),
        # ✅ por role + accessible name
        page.get_by_role("textbox", name=re.compile(r"CPF|CNPJ|CNO|CAEPF", re.I)),
    ]

    for c in candidatos:
        if _fill_first(page, c, digits):
            return
    raise RuntimeError("Campo 'CPF/CNPJ/CNO/CAEPF' não encontrado.")

def fill_cnae_primario(page, codigo="00.00-0-00"):
    """Preenche CNAE e seleciona a primeira opção do autocomplete."""
    # Campo é um combobox com autocomplete (mat-autocomplete)
    loc = page.get_by_placeholder(re.compile(r"CNAE\s*Prim[aá]rio", re.I))
    if not loc.count():
        loc = page.locator('input[role="combobox"][placeholder*="CNAE"]')
    if not loc.count():
        raise RuntimeError("Campo 'CNAE Primário' não encontrado.")

    caixa = loc.first
    caixa.scroll_into_view_if_needed()
    caixa.click()
    caixa.fill(codigo)
    caixa.press("Enter")

    # Aguarda a abertura do listbox e seleciona a 1ª opção
    try:
        listbox = page.get_by_role("listbox")
        if listbox.count():
            listbox.first.wait_for(state="visible", timeout=10000)
            caixa.press("Enter")
            # Espera fechar
            try:
                listbox.first.wait_for(state="hidden", timeout=5000)
            except PWTimeout:
                pass
    except PWTimeout:
        pass

    # Garante que o combobox não está mais expandido
    try:
        input_id = caixa.get_attribute("id")
        if input_id:
            page.wait_for_function(
                "id => (document.getElementById(id)?.getAttribute('aria-expanded')) === 'false'",
                arg=input_id,
                timeout=5000,
            )
    except PWTimeout:
        pass


# =====================================
# Modal de Regime Tributário
# =====================================

def abrir_modal_regime(page):
    # Botão com tooltip que menciona Regime Tributário
    btn = page.locator('button[mattooltip*="Regime Tributário"]').first
    if not btn.count():
        # Fallback por texto do SVG/ícone caneta
        btn = page.locator("button").filter(has_text=re.compile(r"Regime|Situa[çc][aã]o", re.I)).first
    if not btn.count():
        raise RuntimeError("Botão para alterar Regime/Situação não encontrado.")
    btn.scroll_into_view_if_needed()
    btn.click()

    # Espera o overlay do diálogo
    overlay = page.locator(".cdk-overlay-pane .mat-mdc-dialog-container").first
    overlay.wait_for(state="visible", timeout=15000)
    return overlay

def set_regime_autonomo_e_criar(page, overlay):
    # Campo "Regime Tributário" dentro do overlay
    campo = overlay.get_by_label(re.compile(r"Regime\s*Tribut[aá]rio", re.I))
    if not campo.count():
        campo = overlay.locator(
            "mat-form-field:has(mat-label:has-text('Regime Tributário')) input"
        )
    if not campo.count():
        raise RuntimeError("Campo 'Regime Tributário' no overlay não encontrado.")

    caixa = campo.first
    caixa.scroll_into_view_if_needed()
    caixa.fill("Autônomo")
    caixa.press("Enter")

    # Se aparecer a lista, seleciona
    try:
        listbox = overlay.get_by_role("listbox")
        if listbox.count():
            listbox.first.wait_for(state="visible", timeout=8000)
            caixa.press("Enter")
            try:
                listbox.first.wait_for(state="hidden", timeout=5000)
            except PWTimeout:
                pass
    except PWTimeout:
        pass

    # Aguarda o ícone mudar de "search" para "close" (critério do usuário)
    try:
        overlay.locator("mat-icon", has_text=re.compile(r"^\s*close\s*$", re.I)).first.wait_for(
            state="visible", timeout=10000
        )
    except PWTimeout:
        # Continua mesmo assim; alguns temas não exibem ícone visível
        pass

    # Clica no botão "Criar" do diálogo
    btn_criar = overlay.get_by_role("button", name=re.compile(r"^\s*Criar\s*$", re.I))
    if not btn_criar.count():
        btn_criar = overlay.locator("button").filter(has_text=re.compile(r"^\s*Criar\s*$", re.I))
    if not btn_criar.count():
        raise RuntimeError("Botão 'Criar' do overlay não encontrado.")

    btn_criar.first.scroll_into_view_if_needed()
    btn_criar.first.click()

    # Espera o overlay fechar
    try:
        overlay.wait_for(state="hidden", timeout=10000)
    except PWTimeout:
        pass

def fill_data_inicio_prestacao_servico(page, data_str="01/01/2025"):
    """
    Preenche o campo 'Data Início da Prestação de Serviço' com data_str.
    Limpa totalmente o campo (seleciona tudo e apaga) antes de digitar.
    Tenta primeiro por placeholder; tem fallbacks por label, id e classe.
    """
    candidatos = [
        # placeholder (regex tolerante a acentos e espaços)
        page.get_by_placeholder(
            re.compile(r"Data\s*In[ií]cio\s+da\s+Presta[çc][ãa]o\s+de\s+Servi[çc]o", re.I)
        ),
        # placeholder por substring (duas variantes por causa do cedilha/til)
        page.locator('[placeholder*="Início da Prestação de Serviço"]'),
        page.locator('[placeholder*="Inicio da Prestacao de Servico"]'),
        # pelo id visto no snippet (fallback direto, pode mudar entre sessões)
        page.locator("#mat-input-22"),
        # qualquer input do mat-datepicker dentro de um mat-form-field com label correspondente
        page.locator(
            'mat-form-field:has(mat-label:has-text("Data Início da Prestação de Serviço")) input.mat-datepicker-input'
        ),
        # fallback bem genérico (último recurso)
        page.locator("input.mat-datepicker-input"),
    ]

    campo = None
    for loc in candidatos:
        try:
            if loc.count():
                campo = loc.first
                campo.wait_for(state="visible", timeout=6000)
                break
        except Exception:
            continue

    if not campo:
        raise RuntimeError("Campo 'Data Início da Prestação de Serviço' não encontrado.")

    # Garante foco
    try:
        campo.scroll_into_view_if_needed()
    except Exception:
        pass
    campo.click()

    # Limpa completamente o valor atual
    # 1) Select All (Windows/Linux)
    try:
        campo.press("Control+A")
    except Exception:
        pass
    # 2) Select All (macOS)
    try:
        campo.press("Meta+A")
    except Exception:
        pass
    # 3) Apaga
    for key in ("Delete", "Backspace"):
        try:
            campo.press(key)
        except Exception:
            pass

    # Algumas máscaras não limpam com Backspace; força vazio
    try:
        campo.fill("")  # fill já limpa e digita; aqui usamos só pra limpar
    except Exception:
        pass

    # Digita devagar pra cooperar com a máscara do Angular Material
    campo.type(data_str, delay=50)

    # Desfoca para disparar validação/formatadores
    try:
        campo.press("Tab")
    except Exception:
        pass

    # Checagem rápida: se o valor não refletiu, tenta abrir o datepicker e fechar (força parse)
    try:
        val = campo.evaluate("el => el.value")
        if not val or len(val.strip()) < 10:
            # procura o toggle do datepicker dentro do mesmo mat-form-field
            root = campo.locator("xpath=ancestor::mat-form-field[1]")
            toggle_btn = root.locator("mat-datepicker-toggle button")
            if toggle_btn.count():
                toggle_btn.first.click()
                # fecha o calendário com Escape (não vamos navegar no calendário)
                page.keyboard.press("Escape")
                # revalida
                campo.click()
                campo.press("Tab")
    except PWTimeout:
        pass

def click_salvar_empresa(page):
    """Clica no botão 'Salvar' na página de empresa e aguarda um estado estável."""
    candidatos = [
        # id fornecido
        page.locator('#btnSaveEmpresa'),
        # por nome acessível
        page.get_by_role("button", name=re.compile(r"^\s*Salvar\s*$", re.I)),
        # botão raised primário contendo o texto
        page.locator('button[mat-raised-button][color="primary"]').filter(
            has_text=re.compile(r"\bSalvar\b", re.I)
        ),
    ]

    for loc in candidatos:
        try:
            if loc.count():
                btn = loc.first
                try:
                    btn.scroll_into_view_if_needed()
                except Exception:
                    pass
                btn.wait_for(state="visible", timeout=8000)

                # aguarda habilitar (validações podem bloquear)
                for _ in range(15):
                    try:
                        if btn.is_enabled():
                            break
                    except Exception:
                        pass
                    time.sleep(0.2)

                btn.click()

                # estabiliza rede/toast
                try:
                    page.wait_for_load_state("networkidle", timeout=8000)
                except PWTimeout:
                    pass

                # snack-bar do Angular Material (se houver)
                try:
                    snack = page.locator('.mat-mdc-snack-bar-container')
                    if snack.count():
                        snack.first.wait_for(state="visible", timeout=5000)
                        try:
                            snack.first.wait_for(state="hidden", timeout=7000)
                        except PWTimeout:
                            pass
                except Exception:
                    pass

                return True
        except Exception:
            continue

    raise RuntimeError("Botão 'Salvar' não encontrado.")
# =====================================
# Planilha Cadastros.xlsx
# =====================================

def carregar_cadastros():
    """
    Lê 'Cadastros.xlsx' na mesma pasta do script.
    Coluna A = nome, Coluna B = CPF. Sem cabeçalho obrigatório.
    Retorna lista de dicts: {"nome": ..., "cpf": ...}
    """
    base = Path(__file__).resolve().parent
    caminho = base / "Cadastros.xlsx"
    if not caminho.exists():
        raise FileNotFoundError("Arquivo 'Cadastros.xlsx' não encontrado na pasta do script.")

    # Lê sem cabeçalho para ser mais tolerante
    df = pd.read_excel(caminho, dtype=str, header=None)
    df = df.fillna("")

    registros = []
    for _, row in df.iterrows():
        nome = str(row.iloc[0]).strip()
        cpf  = str(row.iloc[1]).strip() if row.shape[0] > 1 else ""
        erp  = str(row.iloc[2])
        if not nome or not cpf or not erp:
            continue
        # Ignora linhas de cabeçalho comuns
        if nome.lower() in {"nome", "name"} or cpf.lower() in {"cpf", "documento"} or erp.lower() in {"erp"}:
            continue
        registros.append({"nome": nome, "cpf": cpf, "erp": erp})

    if not registros:
        raise ValueError("Nenhuma linha válida encontrada (precisa de Nome na coluna A, CPF na coluna B e ERP na coluna C).")
    return registros

# =====================================
# Fluxo principal
# =====================================

def criar_empresa(page, nome, cpf, erp):
    """Abre o form e executa todo o preenchimento requerido."""
    open_nova_empresa(page)

    fill_codigo_erp(page, erp)
    fill_nome_fantasia(page, nome)
    fill_razao_social(page, nome)
    fill_cpf_cnpj(page, cpf)

    fill_cnae_primario(page, "00.00-0-00")
    
    fill_data_inicio_prestacao_servico(page, "01/01/2025")
    
    overlay = abrir_modal_regime(page)
    set_regime_autonomo_e_criar(page, overlay)
    click_salvar_empresa(page)

    # Pequena pausa para toasts/eventos assíncronos
    time.sleep(1.5)

def fluxo():
    registros = carregar_cadastros()

    with sync_playwright() as p:
        # Tenta usar Chrome; se indisponível, cai para Chromium embarcado
        try:
            navegador = p.chromium.launch(channel="chrome", headless=False)
        except Exception:
            navegador = p.chromium.launch(headless=False)

        contexto = navegador.new_context(
            accept_downloads=True,
            viewport={"width": 1440, "height": 900},
        )
        pagina = contexto.new_page()
        pagina.set_default_timeout(45000)

        # Abre módulo de empresas (assume sessão já autenticada)
        pagina.goto(URL)
        wait_dom(pagina)
        
        fill_email(pagina, "contabil32@wilsonlopes.com.br")
        fill_password(pagina, "123456")
        click_entrar(pagina)
        wait_dom(pagina)

        # 2) Overlay e vai direto para /servicos/novo
        click_overlay_submit(pagina)

        # Espera determinística pela troca de URL para /servicos_programados
        pagina.wait_for_url(
            re.compile(r"^https://web\.tareffa\.com\.br/servicos_programados(?:/)?(?:\?.*)?(?:#.*)?$"),
            timeout=45000
        )
        wait_dom(pagina)  # garante domcontentloaded

        open_nova_empresa(pagina)

        # Processa cada linha da planilha
        for i, item in enumerate(registros, start=1):
            nome = item["nome"]
            cpf  = item["cpf"]
            erp  = item["erp"]
            print(f"[{i}/{len(registros)}] Criando empresa para: {nome} - CPF: {cpf}")
            criar_empresa(pagina, nome, cpf, erp)

        print("✅ Todos os cadastros foram processados.")

if __name__ == "__main__":
    fluxo()
