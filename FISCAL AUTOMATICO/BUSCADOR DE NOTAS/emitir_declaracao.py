import os
import re
import sys
import time
import unicodedata
import subprocess
from datetime import datetime

from playwright.sync_api import sync_playwright
import shutil

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _stdin_interativo() -> bool:
    try:
        return bool(sys.stdin) and sys.stdin.isatty()
    except Exception:
        return False


def _input_seguro(prompt: str, default: str = "") -> str:
    try:
        return input(prompt)
    except EOFError:
        if default:
            print(f"\nEntrada indisponível. Usando padrão: {default}")
        else:
            print("\nEntrada indisponível. Seguindo com valor vazio.")
        return default


def fechar_modais_bloqueantes(pagina, timeout_segundos: float = 6.0) -> None:
    fim = time.time() + timeout_segundos
    botoes_fechamento = [
        "#avisosDoSistemaModal button[data-dismiss='modal']",
        "#avisosDoSistemaModal button:has-text('Fechar')",
        "#avisosDoSistemaModal button:has-text('OK')",
        "#avisosDoSistemaModal button:has-text('Ok')",
        "#avisosDoSistemaModal button:has-text('Entendi')",
        "div.modal.in button[data-dismiss='modal']",
        "div.modal.show button[data-dismiss='modal']",
    ]

    while time.time() < fim:
        houve_acao = False

        for seletor in botoes_fechamento:
            try:
                botao = pagina.locator(seletor).first
                if botao.count() > 0 and botao.is_visible():
                    botao.click(timeout=800)
                    houve_acao = True
                    time.sleep(0.15)
            except Exception:
                pass

        # Limpeza defensiva para backdrop que continua bloqueando clique.
        try:
            pagina.evaluate(
                """
                () => {
                    document.querySelectorAll('.modal-backdrop').forEach((el) => el.remove());
                    document.querySelectorAll('.modal.in, .modal.show').forEach((m) => {
                        m.classList.remove('in', 'show');
                        m.style.display = 'none';
                        m.setAttribute('aria-hidden', 'true');
                    });
                    document.body.classList.remove('modal-open');
                }
                """
            )
        except Exception:
            pass

        try:
            modais_visiveis = pagina.locator("div.modal.in, div.modal.show").count()
            backdrops = pagina.locator("div.modal-backdrop").count()
            if modais_visiveis == 0 and backdrops == 0:
                return
        except Exception:
            pass

        if not houve_acao:
            time.sleep(0.2)


def emitir_declaracoes_disponiveis(pagina, nome_prestador, mes, ano):
    """
    Localiza os botões de emitir declaração na tela atual e emite TODAS
    as declarações disponíveis para o prestador / competência informados.
    (já estando na tela DECLARAÇÃO > Sem movimento)
    """
    wrapper_sel = (
        "div#page-wrapper div.panel.panel-primary div.panel-body "
        "div#tabelaDinamica_wrapper.dataTables_wrapper.form-inline.dt-bootstrap.no-footer"
    )

    # Garante que a tabela carregou
    try:
        pagina.wait_for_selector(wrapper_sel, timeout=8000)
    except Exception:
        print(f"[{nome_prestador}] Não encontrei a tabela de declarações em {mes:02d}/{ano}.")
        return

    # Coleta os IDs de emitirDeclaracao(...)
    botoes_disponiveis = []
    try:
        links = pagina.locator(f'{wrapper_sel} a[href*="emitirDeclaracao"]')
        total = links.count()
        for i in range(total):
            href = links.nth(i).get_attribute("href") or ""
            m = re.search(r"emitirDeclaracao\('(\d+)'\)", href)
            if m:
                botoes_disponiveis.append(int(m.group(1)))
    except Exception as e:
        print(f"[{nome_prestador}] Erro ao coletar botões de declaração: {e}")
        return

    if not botoes_disponiveis:
        print(f"[{nome_prestador}] Nenhuma declaração disponível para {mes:02d}/{ano}.")
        return

    # Emite TODAS as declarações encontradas
    for numero_mes in sorted(set(botoes_disponiveis)):
        try:
            pagina.evaluate(f"emitirDeclaracao('{numero_mes}')")
            pagina.wait_for_load_state("domcontentloaded")
            time.sleep(0.5)

            pagina.click("text=Gravar", timeout=30000)
            pagina.wait_for_load_state("domcontentloaded")

            print(f"[{nome_prestador}] Declaração sem movimento GERADA {numero_mes:02d}/{ano}.")
        except Exception as e:
            print(f"[{nome_prestador}] ERRO ao gravar declaração {numero_mes:02d}/{ano}: {e}")

        # Volta para a tela de pesquisa de declarações para o próximo número
        try:
            pagina.click("text=Pesquisar")
        except Exception:
            pass
        pagina.wait_for_load_state("domcontentloaded")
        time.sleep(0.5)


def _normalizar(txt: str) -> str:
    return unicodedata.normalize("NFKD", txt or "").encode("ASCII", "ignore").decode("ASCII").lower().strip()


def _extrair_cnpj(texto_opcao: str) -> str | None:
    # Tenta achar 14 dígitos seguidos
    limpo = re.sub(r"\D", "", texto_opcao or "")
    m = re.search(r"(\d{14})", limpo)
    if m:
        return m.group(1)
    return None


# Mês/ano de referência = mês anterior
hoje = datetime.today()
mes_anterior = hoje.month - 1 or 12
ano_ref = hoje.year if hoje.month > 1 else hoje.year - 1
meses_ext = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
]
mes_extenso = meses_ext[mes_anterior - 1]

def escolher_perfil_firefox(perfil_padrao: str) -> tuple[str, bool]:
    """
    Pergunta se deve usar um perfil existente ou criar um novo.

    Retorna (caminho_perfil, perfil_resetado).
    - perfil_resetado=True => o perfil_padrao foi APAGADO e recriado vazio.
      Também é recomendado forçar a seleção de certificado (Ask Every Time).
    """
    # Em execuções sem terminal interativo (agendador/serviço), não pergunta.
    if not _stdin_interativo():
        os.makedirs(perfil_padrao, exist_ok=True)
        return perfil_padrao, False

    print("\n=== PERFIL FIREFOX ===")
    print("1) Usar perfil existente (mantém certificado previamente selecionado)")
    print("2) Criar NOVO perfil (apaga o perfil antigo e força seleção de certificado)\n")
    opcao = (_input_seguro("Escolha [1/2] (padrão=1): ", default="1").strip() or "1")

    if opcao == "2":
        # Apaga o perfil antigo (o que fica salvo na pasta padrão)
        try:
            if os.path.isdir(perfil_padrao):
                shutil.rmtree(perfil_padrao)
        except Exception as e:
            print(f"Não foi possível apagar o perfil antigo em '{perfil_padrao}': {e}")
            print("Feche qualquer Firefox usando esse perfil e tente novamente.")
            raise
        os.makedirs(perfil_padrao, exist_ok=True)
        return perfil_padrao, True

    if opcao == "1":
        caminho = _input_seguro(
            f"Caminho do perfil (Enter para usar o padrão: {perfil_padrao}): ",
            default=perfil_padrao,
        ).strip()
        caminho = caminho or perfil_padrao
        os.makedirs(caminho, exist_ok=True)
        return caminho, False

    print("Opção inválida; usando o perfil padrão.")
    os.makedirs(perfil_padrao, exist_ok=True)
    return perfil_padrao, False

# ================== FLUXO PRINCIPAL ==================
with sync_playwright() as p:
    # Perfil exclusivo (persistente) da Playwright / Firefox
    perfil_padrao = r"C:\ROBOS\perfis firefox\BUSCA NOTAS\perfil_firefox_cert_esnfs"
    perfil_firefox, perfil_resetado = escolher_perfil_firefox(perfil_padrao)

    # Preferências ligadas ao certificado:
    # - Usando perfil existente: mantém o comportamento atual (seleção automática / certificado já escolhido)
    # - Perfil novo (resetado): força a tela de seleção ("Ask Every Time") e evita 'lembrar' automaticamente
    firefox_prefs = {
        "security.default_personal_cert": "Select Automatically",
        "security.remember_cert_checkbox_default_setting": True,
    }
    if perfil_resetado:
        firefox_prefs.update({
            "security.default_personal_cert": "Ask Every Time",
            "security.remember_cert_checkbox_default_setting": False,
        })

    contexto = p.firefox.launch_persistent_context(
        user_data_dir=perfil_firefox,
        headless=True,
        accept_downloads=False,
        firefox_user_prefs=firefox_prefs,
    )
    pagina = contexto.new_page()

    # Acesso ao site
    pagina.goto("https://www.esnfs.com.br/?e=35")
    time.sleep(3)

    # Fecha modais de aviso que possam bloquear a tela inicial
    fechar_modais_bloqueantes(pagina)


    # Login por CERTIFICADO DIGITAL
    botao_cert = pagina.locator('button[onclick*="useDigitalCertificate=true"]')
    botao_cert.wait_for(state="visible", timeout=1000)
    botao_cert.click()
    fechar_modais_bloqueantes(pagina)


    # Seleciona o município
    pagina.wait_for_selector("text=Município de Francisco Beltrão", timeout=30000)
    pagina.click("text=Município de Francisco Beltrão")
    time.sleep(3)

    # Alguns segundos para ver se logou mesmo
    url_atual = pagina.url
    if (
        "login" in url_atual.lower()
        or "captcha" in url_atual.lower()
        or pagina.locator("iframe[title*='recaptcha']").count() > 0
    ):
        print("🚫 reCAPTCHA bloqueou o login — reiniciando script...")
        contexto.close()
        time.sleep(3)
        subprocess.run([sys.executable, sys.argv[0]], check=True)
        sys.exit(0)

    print("✅ Login efetuado com sucesso, prosseguindo...")

    # Garante navegação para a tela correta antes de buscar selects de exercício/prestador.
    pagina.goto("https://www.esnfs.com.br/nfsdeclaracao.list.logic")
    pagina.wait_for_load_state("domcontentloaded")
    fechar_modais_bloqueantes(pagina)

    # Preenche o ano
    pagina.wait_for_selector('select[name="formulario.nrExercicio"]', timeout=60000)
    ano_select = pagina.locator('select[name="formulario.nrExercicio"]')
    ano_select.select_option(value=str(ano_ref))

    # Coleta lista de prestadores a partir da própria tela de DECLARAÇÃO
    pagina.wait_for_selector('select[name="formulario.pessoa.idPessoa"]', timeout=60000)
    prestador_select = pagina.locator('select[name="formulario.pessoa.idPessoa"]')
    prestadores = prestador_select.locator("option").all()
    prestadores_info = [
        (p.get_attribute("value"), p.text_content().strip())
        for p in prestadores
        if p.get_attribute("value")
    ]

    # Mapa de índices por "value" e por CNPJ
    indice_por_valor = {}
    indice_por_cnpj = {}
    for idx, (valor, texto) in enumerate(prestadores_info):
        indice_por_valor[valor] = idx
        cnpj = _extrair_cnpj(texto)
        if cnpj:
            indice_por_cnpj[cnpj] = idx

    # Escolha do ponto inicial
    print(
        "\n➤ Você pode informar:\n"
        "  • CNPJ (com ou sem máscara)\n"
        "  • Código 'value' (do select)\n"
        "  • Parte do nome da empresa\n"
        "  • Ou Enter para iniciar a partir do ID 0 (se existir).\n"
    )

    if _stdin_interativo():
        entrada = _input_seguro("Digite aqui qual empresa deseja iniciar (ou Enter): ", default="").strip()
    else:
        entrada = ""
    if entrada:
        entrada_digits = re.sub(r"\D", "", entrada)
        indice_inicio = None

        # 1) value
        if entrada in indice_por_valor:
            indice_inicio = indice_por_valor[entrada]
            print(f"➡️ Iniciando pelo código (value) '{entrada}'.")

        # 2) CNPJ
        if indice_inicio is None and len(entrada_digits) == 14 and entrada_digits in indice_por_cnpj:
            indice_inicio = indice_por_cnpj[entrada_digits]
            print(f"➡️ Iniciando pelo CNPJ {entrada_digits}.")

        # 3) parte do nome
        if indice_inicio is None:
            alvo = _normalizar(entrada)
            for i, (_, texto) in enumerate(prestadores_info):
                if alvo and _normalizar(texto).find(alvo) != -1:
                    indice_inicio = i
                    print(f"➡️ Iniciando pelo nome que contém: '{entrada}'.")
                    break

        # fallback
        if indice_inicio is None:
            print("⚠️ Não encontrei essa empresa. Iniciando a partir do ID 0 (quando existir).")
            indice_inicio = next(
                (i for i, (valor, _) in enumerate(prestadores_info) if valor == "0"), 0
            )
    else:
        indice_inicio = next(
            (i for i, (valor, _) in enumerate(prestadores_info) if valor == "0"), 0
        )
        print("↩️ Nenhuma escolha informada. Iniciando a partir do ID 0 (quando existir).")

    prestadores_a_processar = prestadores_info[indice_inicio:]
    if not prestadores_a_processar:
        raise Exception(
            "❌ Nenhum prestador válido encontrado para processar a partir do índice escolhido."
        )


    # ===================== LOOP PRINCIPAL (DECLARAÇÕES) =====================
    pagina.goto("https://www.esnfs.com.br/nfsdeclaracao.list.logic")

    indices_erro_msg = []


    for idx, (valor_prestador, nome_prestador_completo) in enumerate(prestadores_a_processar):
        print(
            f"\n🔍 Processando prestador (DECLARAÇÃO): "
            f"{nome_prestador_completo} (valor: {valor_prestador})"
        )

        # Limpa filtros na PRÓPRIA tela de declaração
        try:
            pagina.click("text=Limpar")
            pagina.wait_for_load_state("domcontentloaded")
        except Exception:
            pass

        # Preenche ano e seleciona prestador e competência (na tela de declaração)
        try:
            ano_select = pagina.locator('select[name="formulario.nrExercicio"]')
            ano_select.select_option(value=str(ano_ref))

            prestador_select = pagina.locator('select[name="formulario.pessoa.idPessoa"]')
            prestador_select.select_option(value=valor_prestador)

            pagina.click("text=Pesquisar")
            pagina.wait_for_load_state("domcontentloaded")

            # Se aparecer a mensagem de erro, marca o índice e pula essa empresa
            try:
                # espera curtinho: se não aparecer, segue normal
                pagina.wait_for_selector("#mensagemErro .alert.alert-danger", state="visible", timeout=1200)
                indices_erro_msg.append(idx)
                print(f"[{nome_prestador_completo}] ⚠️ Aviso de erro detectado (#mensagemErro). Vai para o próximo e agenda retry.")
                
                pagina.goto("https://www.esnfs.com.br/nfsdeclaracao.list.logic")
                continue
            except Exception:
                pass  # não apareceu mensagem -> segue o fluxo normal


            emitir_declaracoes_disponiveis(
                pagina=pagina,
                nome_prestador=nome_prestador_completo,
                mes=mes_anterior,
                ano=ano_ref,
            )

        except Exception as e:
            print(
                f"[{nome_prestador_completo}] ERRO ao emitir declarações "
                f"{mes_anterior:02d}/{ano_ref}: {e}"
            )

    if indices_erro_msg:
        print("\n🔁 Tentando novamente apenas as empresas com aviso de erro...")
        prestadores_retry = [prestadores_a_processar[i] for i in indices_erro_msg]

        # opcional: guardar quem ainda falhar no retry
        indices_erro_msg_retry = []

        for idx_retry, (valor_prestador, nome_prestador_completo) in zip(indices_erro_msg, prestadores_retry):
            print(f"\n🔄 Retry prestador (idx {idx_retry}): {nome_prestador_completo} (valor: {valor_prestador})")

            try:
                pagina.click("text=Limpar")
                pagina.wait_for_load_state("domcontentloaded")
            except Exception:
                pass

            try:
                ano_select = pagina.locator('select[name="formulario.nrExercicio"]')
                ano_select.select_option(value=str(ano_ref))

                prestador_select = pagina.locator('select[name="formulario.pessoa.idPessoa"]')
                prestador_select.select_option(value=valor_prestador)

                pagina.click("text=Pesquisar")
                pagina.wait_for_load_state("domcontentloaded")

                # checa de novo a mesma mensagem de erro
                try:
                    pagina.wait_for_selector("#mensagemErro .alert.alert-danger", state="visible", timeout=1200)
                    indices_erro_msg_retry.append(idx_retry)
                    print(f"[{nome_prestador_completo}] ⚠️ Ainda com aviso de erro no retry.")
                    continue
                except Exception:
                    pass

                emitir_declaracoes_disponiveis(
                    pagina=pagina,
                    nome_prestador=nome_prestador_completo,
                    mes=mes_anterior,
                    ano=ano_ref,
                )

            except Exception as e:
                print(f"[{nome_prestador_completo}] ERRO no retry {mes_anterior:02d}/{ano_ref}: {e}")

        if indices_erro_msg_retry:
            print(f"\n⚠️ Permaneceram com aviso de erro após retry (índices): {indices_erro_msg_retry}")

    print("\n✅ Fim do processamento de declarações sem movimento.")
