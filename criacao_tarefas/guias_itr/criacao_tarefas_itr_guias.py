from logging import PlaceHolder
import os
import re
import glob
import time
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from playwright.sync_api import expect

URL = "https://web.tareffa.com.br/servicos"
NEW_URL = "https://web.tareffa.com.br/servicos/novo"

# =========================
# Helpers de espera
# =========================
def wait_dom(page):
    page.wait_for_load_state("domcontentloaded")

def wait_new_service_form_ready(page):
    """
    Garante que o formulário de 'Novo Serviço' foi renderizado
    (sem time.sleep). Procura por campos ligados à 'Descrição'.
    """
    seletor = (
        'input[aria-describedby*="help-descricao"], '
        'textarea[aria-describedby*="help-descricao"], '
        'mat-form-field:has(mat-label:has-text("Descrição")) input, '
        'mat-form-field:has(mat-label:has-text("Descrição")) textarea'
    )
    page.wait_for_selector(seletor, timeout=45000)

def open_novo_servico(page):
    """
    Abre diretamente o formulário de 'Novo Serviço' e espera o form renderizar.
    """
    page.goto(NEW_URL)
    wait_dom(page)
    wait_new_service_form_ready(page)

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

# =========================
# Formulário de Novo Serviço
# =========================
def fill_service_name(page, descricao):
    # Mais robusto: tenta vários caminhos (input/textarea/label)
    candidatos = [
        page.locator('input[aria-describedby*="help-descricao"]'),
        page.locator('textarea[aria-describedby*="help-descricao"]'),
        page.locator('mat-form-field:has(mat-label:has-text("Descrição")) input'),
        page.locator('mat-form-field:has(mat-label:has-text("Descrição")) textarea'),
        page.get_by_placeholder(re.compile(r"descri", re.I)),
        page.get_by_role("textbox", name=re.compile(r"descri", re.I)),
        page.locator('input[id^="mat-input-"]'),
        page.locator('textarea[id^="mat-input-"]'),
    ]
    for loc in candidatos:
        if loc.count():
            try:
                loc.first.scroll_into_view_if_needed()
            except Exception:
                pass
            try:
                loc.first.wait_for(state="visible", timeout=3000)
                loc.first.fill(descricao)
                return
            except Exception:
                continue
    raise RuntimeError("Campo de descrição/nome do serviço não encontrado.")

def assunto(page, assunto):
    # Mais robusto: tenta vários caminhos (input/textarea/label)
    candidatos = [
        page.locator('input[aria-describedby*="eSubjectHelp"]'),
        page.locator('textarea[aria-describedby*="eSubjectHelp"]'),
    ]
    for loc in candidatos:
        if loc.count():
            try:
                loc.first.scroll_into_view_if_needed()
            except Exception:
                pass
            try:
                loc.first.wait_for(state="visible", timeout=3000)
                loc.first.fill(assunto)
                return
            except Exception:
                continue
    raise RuntimeError("Campo de assunto não encontrado.")

def mensagem(page, texto_msg):
    # candidatos em ordem de “força”
    candidatos = [
        page.locator('textarea[id^="mat-input-"], input[id^="mat-input-"]'),
        page.locator('[contenteditable="true"]'),
    ]

    valor = f"Segue em anexo {texto_msg}."

    for loc in candidatos:
        try:
            # garante que existe pelo menos 1 match no DOM
            if loc.count() == 0:
                continue

            alvo = loc.first
            try:
                alvo.scroll_into_view_if_needed(timeout=2000)
            except PWTimeout:
                pass

            # alguns elementos só ficam 'visible' após scroll
            try:
                alvo.wait_for(state="visible", timeout=3000)
            except PWTimeout:
                # se não ficar visível, ainda tentamos um click forçado
                try:
                    alvo.click(force=True, timeout=1000)
                except Exception:
                    continue

            # contenteditable precisa de tratamento diferente
            ce_attr = alvo.get_attribute("contenteditable")
            if ce_attr and ce_attr.lower() == "true":
                try:
                    alvo.click(timeout=2000)
                except Exception:
                    alvo.click(force=True, timeout=2000)
                page.keyboard.insert_text(valor)
            else:
                # tenta fill; se não rolar, usa type
                try:
                    alvo.fill(valor, timeout=3000)
                except Exception:
                    try:
                        alvo.click(timeout=2000)
                    except Exception:
                        alvo.click(force=True, timeout=2000)
                    # limpa com Ctrl/Command+A e digita
                    with page.expect_event("console", predicate=lambda m: False, timeout=100): 
                        # no-op para não bloquear
                        pass
                    page.keyboard.press("Control+A")
                    page.keyboard.type(valor, delay=10)

                # dispara eventos para Angular/Material detectar a mudança
                page.evaluate("""
                el => {
                    el.dispatchEvent(new Event('input',  {bubbles:true}));
                    el.dispatchEvent(new Event('change', {bubbles:true}));
                    el.dispatchEvent(new Event('blur',   {bubbles:true}));
                }
                """, alvo)

            return  # sucesso: sai da função

        except Exception:
            # tenta o próximo candidato
            continue

    raise RuntimeError("Campo de mensagem não encontrado.")

def select_departamento_itr(page):
    # Autocomplete: placeholder 'Tecle Enter para pesquisar um departamento'
    dep = page.get_by_placeholder(re.compile(r"pesquisar.*departamento", re.I))
    if not dep.count():
        dep = page.locator('input[aria-describedby*="help-departamento"]')
    if not dep.count():
        dep = page.locator("#mat-input-4")
    if not dep.count():
        raise RuntimeError("Campo de departamento não encontrado.")

    caixa = dep.first
    caixa.wait_for(state="visible")
    caixa.fill("Departamento ITR")
    caixa.press("Enter")

    # Se abrir lista (role=listbox), mais um Enter para selecionar a 1ª opção
    try:
        listbox = page.get_by_role("listbox")
        if listbox.count():
            listbox.first.wait_for(state="visible", timeout=5000)
            caixa.press("Enter")
    except PWTimeout:
        pass

def marcar_checkbox_baixa_manual(page):
    chk = page.locator('#permiteBaixaManual-input')
    if not chk.count():
        chk = page.locator('input.mdc-checkbox__native-control#permiteBaixaManual-input')
    if not chk.count():
        raise RuntimeError("Checkbox 'permiteBaixaManual-input' não encontrada.")
    try:
        chk.first.check()
    except Exception:
        chk.first.click()
        
def marcar_checkbox_comunicacao(page):
    # ---- 1) Comunica Cliente quando Concluído ----
    input1 = page.locator("#checkComunicaCliente")
    label1 = page.locator('label[for="checkComunicaCliente"]')

    # Garante que estão no DOM
    input1.wait_for(state="attached", timeout=1500)
    label1.wait_for(state="visible", timeout=1500)

    try:
        # Tenta a forma "oficial" (pode falhar por overlay do label)
        input1.check(timeout=1500)
    except Exception:
        try:
            # Fallback mais natural em custom-checkbox: clicar no label
            label1.scroll_into_view_if_needed()
            label1.click(timeout=1500)
        except PWTimeout:
            # Último recurso: forçar a marcação (ignora actionability)
            input1.check(force=True, timeout=1500)

    # Sanity check da 1ª
    if not input1.is_checked():
        raise RuntimeError("Falha ao marcar 'Comunica Cliente quando Concluído'.")

    # ---- 2) Encerra comunicação após sua abertura ----
    input2 = page.locator("#checkEncerraComunicacao")
    label2 = page.locator('label[for="checkEncerraComunicacao"]')

    # Está disabled no HTML inicialmente; aguarde ficar enabled
    input2.wait_for(state="attached", timeout=2000)
    try:
        # espere até habilitar (o expect faz polling de enabled/disabled)
        expect(input2).to_be_enabled(timeout=3000)
    except Exception:
        # Se nunca habilitar, informe claramente o motivo
        raise RuntimeError("A checkbox 'Encerra comunicação...' permaneceu desabilitada; "
                           "verifique se a 1ª ativa esta opção ou se há validação pendente.")

    # Com enabled, tente marcar
    try:
        input2.check(timeout=1500)
    except Exception:
        try:
            label2.scroll_into_view_if_needed()
            label2.click(timeout=3000)
        except PWTimeout:
            input2.check(force=True, timeout=3000)

    # Sanity check da 2ª
    if not input2.is_checked():
        raise RuntimeError("Falha ao marcar 'Encerra comunicação após sua abertura'.")

def abrir_e_escolher_perfil(page, texto_label="Perfil principal para baixa de documentos", opcao="Financeiro"):
    # 1) Localiza o mat-select associado ao mat-label pelo texto visível
    wrapper = page.locator(
        'mat-form-field:has(mat-label:has-text("{}"))'.format(texto_label)
    )
    # garante que o form-field está no DOM
    wrapper.first.wait_for(state="attached", timeout=10000)

    mat_select = wrapper.first.locator("mat-select").first
    # às vezes o mat-select não está 'visible' ainda, mas attached; scroll ajuda
    try:
        mat_select.scroll_into_view_if_needed(timeout=5000)
    except PWTimeout:
        # se não scrolar, ainda tentamos clicar; se falhar, levantamos erro mais claro
        pass

    # 2) Abre o select
    try:
        mat_select.click(timeout=5000)
    except PWTimeout:
        # último recurso: força o click (em apps com overlay/animations pode ser necessário)
        mat_select.click(timeout=3000, force=True)

    # 3) Descobre o panel via aria-controls OU usa overlay genérico
    panel_id = mat_select.get_attribute("aria-controls")
    if panel_id:
        page.wait_for_selector(f"#{panel_id}", state="visible", timeout=5000)
        options_scope = page.locator(f"#{panel_id}")
    else:
        page.wait_for_selector(".cdk-overlay-pane mat-option", state="visible", timeout=5000)
        options_scope = page.locator(".cdk-overlay-pane")

    # 4) Clica na opção pelo texto (exato, com trims)
    options_scope.locator("mat-option").filter(
        has_text=re.compile(r"^\s*{}\s*$".format(re.escape(opcao)), re.I)
    ).first.click()

    # 5) Espera o painel fechar e o valor refletir
    select_id = mat_select.get_attribute("id")
    if panel_id:
        page.wait_for_selector(f"#{panel_id}", state="hidden", timeout=5000)

    if select_id:
        # aria-expanded deve ficar 'false'
        page.wait_for_function(
            "(sel) => document.querySelector(sel)?.getAttribute('aria-expanded') === 'false'",
            arg=f"#{select_id}",
            timeout=5000
        )

    # 6) Confirma que o valor escolhido aparece no trigger
    wrapper.first.locator(".mat-mdc-select-value").filter(
        has_text=re.compile(re.escape(opcao), re.I)
    ).first.wait_for(state="visible", timeout=5000)

# =========================
# Planilha LISTA ITRS
# =========================
def carregar_lista_itrs():
    """
    Procura por 'LISTA ITRS.*' na pasta do script.
    Suporta .xlsx, .xls, .csv.
    Retorna lista de dicts: {"nome": ..., "cib": ...}
    """
    base = Path(__file__).resolve().parent
    candidatos = []
    for ext in ("xlsx", "xls", "csv"):
        candidatos.extend(glob.glob(str(base / f"LISTA ITRS*.{ext}")))

    if not candidatos:
        raise FileNotFoundError("Arquivo 'LISTA ITRS' não encontrado na pasta do script.")

    caminho = sorted(candidatos)[0]
    ext = Path(caminho).suffix.lower()

    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(caminho, dtype=str, header=0)
    elif ext == ".csv":
        try:
            df = pd.read_csv(caminho, dtype=str, header=0)
        except Exception:
            df = pd.read_csv(caminho, dtype=str, header=0, sep=";")
    else:
        raise ValueError(f"Extensão não suportada: {ext}")

    if df.shape[1] < 4:
        raise ValueError("Planilha 'LISTA ITRS' precisa ter pelo menos 4 colunas (A=Nome, D=CIB).")

    df = df.fillna("")
    registros = []
    for _, row in df.iterrows():
        nome = str(row.iloc[0]).strip()
        cib  = str(row.iloc[3]).strip()
        if nome and cib:
            registros.append({"nome": nome, "cib": cib})
    if not registros:
        raise ValueError("Nenhuma linha válida encontrada (precisa de Nome na coluna A e CIB na coluna D).")
    return registros

# =========================
# Fluxo principal
# =========================
def fluxo():
    registros = carregar_lista_itrs()

    with sync_playwright() as p:
        navegador = p.chromium.launch(channel="chrome", headless=False)
        contexto = navegador.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080},
        )
        pagina = contexto.new_page()
        pagina.set_default_timeout(45000)

        # 1) Abre app e login
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

        open_novo_servico(pagina)

        # 3) Para cada registro, cria um novo serviço indo direto ao form
        for i, item in enumerate(registros, start=1):
            nome_servico = f"DARF ITR {item['nome']} {item['cib']}"
            print(f"[{i}/{len(registros)}] Criando serviço: {nome_servico}")

            # Garante que estamos no form /servicos/novo e que o formulário está pronto
            open_novo_servico(pagina)

            # Preenche formulário
            fill_service_name(pagina, nome_servico)
            select_departamento_itr(pagina)   # usa 'Departamento ITR' conforme sua função
            marcar_checkbox_baixa_manual(pagina)

            # ---- NOVO TRECHO: ir para a aba Programação e preencher ----

            # 1) Ir para a aba "Programação"
            aba_prog = pagina.get_by_role("tab", name=re.compile(r"Programação", re.I))
            if aba_prog.count():
                aba_prog.first.click()
            else:
                pagina.locator('.mat-mdc-tab-labels .mdc-tab .mdc-tab__text-label', has_text="Programação").first.click()

            # aguarda os campos da aba aparecerem (sem sleep)
            pagina.wait_for_selector('input[type="number"][min="1"][max="31"], input[placeholder="01"][type="number"]', timeout=45000)

            # 2) Preencher o dia "30"
            dia = pagina.locator('input[type="number"][min="1"][max="31"]')
            if not dia.count():
                dia = pagina.locator('input[placeholder="01"][type="number"]')
            dia.first.scroll_into_view_if_needed()
            dia.first.click()
            dia.first.fill("30")

            # 3) Marcar "Setembro" – versão rápida
            alvo_lbl = pagina.locator('label.custom-control-label[for="check-set"]').first
            if not alvo_lbl.count():
                alvo_lbl = pagina.locator('label.custom-control-label', has_text=re.compile(r'\bSetembro\b', re.I)).first

            alvo_lbl.scroll_into_view_if_needed()
            alvo_lbl.click()

            # verificação rápida (sem waits longos)
            marcado = False
            try:
                marcado = pagina.locator('#check-set').evaluate("el => !!el && el.checked")
            except Exception:
                pass

            if not marcado:
                # fallback único no input (pode estar invisível, então force=True)
                inp = pagina.locator('#check-set').first
                if inp.count():
                    try:
                        inp.check(force=True, timeout=500)
                    except Exception:
                        # última tentativa: clicar o label mais uma vez
                        alvo_lbl.click()

            # confirmação com timeout zero (sincrono)
            assert pagina.locator('#check-set').evaluate("el => !!el && el.checked"), "Falha ao marcar 'Setembro'"


            # 4) Em "Competência", selecionar "Mesmo Mês do Vencimento"

            # localiza e abre o <mat-select> de Competência
            mat_select = pagina.locator("mat-form-field", has_text=re.compile(r"Compet[êe]ncia", re.I)).locator("mat-select").first
            mat_select.scroll_into_view_if_needed()
            mat_select.click()

            # descobre o painel do overlay via aria-controls (ex.: 'mat-select-100-panel')
            panel_id = mat_select.get_attribute("aria-controls")
            if panel_id:
                pagina.wait_for_selector(f"#{panel_id}", state="visible", timeout=5000)
                options_scope = pagina.locator(f"#{panel_id}")
            else:
                # fallback: usa o overlay genérico do CDK
                pagina.wait_for_selector(".cdk-overlay-pane mat-option", state="visible", timeout=5000)
                options_scope = pagina.locator(".cdk-overlay-pane")

            # clica na opção pelo texto
            options_scope.locator("mat-option").filter(
                has_text=re.compile(r"^\s*Mesmo M[eê]s do Vencimento\s*$", re.I)
            ).first.click()

            # espera o select fechar e o valor refletir a escolha
            select_id = mat_select.get_attribute("id")
            if panel_id:
                pagina.wait_for_selector(f"#{panel_id}", state="hidden", timeout=5000)
            if select_id:
                pagina.wait_for_function(
                    "(sel) => document.querySelector(sel)?.getAttribute('aria-expanded') === 'false'",
                    arg=f"#{select_id}",
                    timeout=5000
                )
            pagina.locator(".mat-mdc-select-value-text").filter(
                has_text=re.compile(r"Mesmo M[eê]s do Vencimento", re.I)
            ).first.wait_for(state="visible", timeout=5000)
            
            # Preencher "Dias de Antecedência (Entrega ao Cliente)" com 5 e dar TAB
            dias = pagina.get_by_placeholder(re.compile(r"Dias de Anteced[eê]ncia\s*\(Entrega ao Cliente\)", re.I))
            if not dias.count():
                dias = pagina.locator('input[type="number"][placeholder*="Dias de Anteced"]')
            dias.first.scroll_into_view_if_needed()
            dias.first.wait_for(state="visible", timeout=5000)
            dias.first.fill("5")
            dias.first.press("Tab")
            
            # 4.5) Ir para a aba "Comunicação"
            aba_prog = pagina.get_by_role("tab", name=re.compile(r"Comunicação", re.I))
            if aba_prog.count():
                aba_prog.first.click()
            else:
                pagina.locator('.mat-mdc-tab-labels .mdc-tab .mdc-tab__text-label', has_text="Comunicação").first.click()
                
            pagina.wait_for_selector('input[type="text"][placeholder*="Assunto..."]', timeout=45000)
                
            marcar_checkbox_comunicacao(pagina)
            
            abrir_e_escolher_perfil(
                pagina,
                texto_label="Perfil principal para baixa de documentos",  # ← texto correto
                opcao="Financeiro"
            )
            
            assunto(pagina, nome_servico)
            mensagem(pagina, nome_servico)

            # 5) Ir para a aba "Integrações"
            aba_prog = pagina.get_by_role("tab", name=re.compile(r"Integrações", re.I))
            if aba_prog.count():
                aba_prog.first.click()
            else:
                pagina.locator('.mat-mdc-tab-labels .mdc-tab .mdc-tab__text-label', has_text="Integrações").first.click()

            # 6) Preencher "Termos para baixa por conteúdo" com CIB formatado e depois o campo numérico

            # Formata o CIB: mantém só dígitos e insere '-' antes do último dígito
            digits = re.sub(r"\D", "", item["cib"])
            if len(digits) >= 2:
                corpo = digits[:-1].zfill(7)   # garante 7 dígitos no corpo
                dv = digits[-1]                # último dígito (verificador)
                cib_fmt = f"{corpo[0]}.{corpo[1:4]}.{corpo[4:]}-{dv}"
            else:
                cib_fmt = digits
            texto_termos = f"1070 ITR - IMPOSTO PROPRIEDADE TERRITORIAL RURAL,CIB:,{cib_fmt}"

            # Escopo da aba ativa (Integrações)
            aba_ativa = pagina.locator(".mat-mdc-tab-body.mat-mdc-tab-body-active")

            # Campo "Termos para baixa por conteúdo"
            termos_input = pagina.get_by_label(re.compile(r"Termos\s+para\s+baixa\s+por\s+conte[úu]do", re.I))
            if not termos_input.count():
                termos_input = aba_ativa.locator(
                    "mat-form-field:has(mat-label:has-text('Termos para baixa por conteúdo')) input, "
                    "mat-form-field:has(mat-label:has-text('Termos para baixa por conteúdo')) textarea"
                )
            if not termos_input.count():
                # usa relação label[for=...] -> input#id
                lbl = aba_ativa.locator("label.mdc-floating-label", has_text=re.compile(r"Termos.*conte[úu]do", re.I)).first
                for_id = lbl.get_attribute("for") if lbl.count() else None
                termos_input = pagina.locator(f"#{for_id}") if for_id else aba_ativa.locator("input[matinput], textarea[matinput]").first

            termos_input.first.scroll_into_view_if_needed()
            termos_input.first.fill(texto_termos)

            # Campo numérico (preencher com "1")
            num_input = aba_ativa.locator("#mat-input-15")
            if not num_input.count():
                num_input = aba_ativa.locator('input[type="number"]').first
            num_input.scroll_into_view_if_needed()
            num_input.fill("2")
            
            # 7) Clicar em "Criar"
            btn_criar = pagina.get_by_role("button", name=re.compile(r"^\s*Criar\s*$", re.I))
            if not btn_criar.count():
                btn_criar = pagina.locator("button").filter(has_text=re.compile(r"^\s*Criar\s*$", re.I))
            btn_criar.first.scroll_into_view_if_needed()
            btn_criar.first.click()

            # Confirma toast de sucesso (robusto a variações)
            time.sleep(3)

        print("✅ Todos os serviços foram processados.")

if __name__ == "__main__":
    fluxo()
