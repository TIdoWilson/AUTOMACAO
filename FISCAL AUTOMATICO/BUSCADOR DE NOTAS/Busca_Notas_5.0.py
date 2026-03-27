import os
import time
from datetime import datetime
import subprocess
from playwright.sync_api import sync_playwright
import unicodedata
import re
import sys
import tempfile

# Caminho base dos downloads
base_download_dir = r"\\192.0.0.251\arquivos\XML PREFEITURA"

# --- LOG de operações (1 linha por operação) ---
LOG_ARQUIVO = None  # resolvido dinamicamente

def _resolver_caminho_log():
    global LOG_ARQUIVO
    try:
        if LOG_ARQUIVO:
            return LOG_ARQUIVO
        base_dir = base_download_dir if 'base_download_dir' in globals() and base_download_dir else os.getcwd()
        LOG_ARQUIVO = os.path.join(base_dir, "log_operacoes.txt")
        return LOG_ARQUIVO
    except Exception:
        LOG_ARQUIVO = os.path.join(os.getcwd(), "log_operacoes.txt")
        return LOG_ARQUIVO

def registrar_log(empresa: str, mensagem: str):
    caminho = _resolver_caminho_log()
    linha = f"{datetime.now():%Y-%m-%d %H:%M:%S} | {empresa} | {mensagem}"
    print(linha)
    try:
        with open(caminho, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception as e:
        print(f"⚠️ Falha ao gravar log: {e}")

# --- util: limpar nome do cliente (sem acentos/ç e sem caracteres especiais) ---
def limpar_nome(nome: str) -> str:
    if not isinstance(nome, str):
        nome = str(nome or "")
        
    # Remove espaços especiais / zero-width
    nome = nome.replace("\xa0", " ")
    nome = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", nome)

    # 1️⃣ Remove acentuação
    nome = unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("ASCII")

    # 2️⃣ Remove prefixo de CNPJ ou CPF seguido de hífen
    nome = re.sub(r'^\s*[\d\.\-\/]{11,18}\s*-\s*', '', nome)

    # 3️⃣ Se houver código final entre parênteses, substitui por espaço e o número (mantém o número)
        #    Ex: "(306681)" -> " 306681"
    nome = re.sub(r'\(\s*(\d+)\s*\)\s*$', r' \1', nome)

    # 4️⃣ Remove qualquer hífen entre palavras
    nome = re.sub(r'\s*[-–—]+\s*', ' ', nome)

    # 5️⃣ Mantém apenas letras, números, espaço, ponto e underscore
    nome = re.sub(r"[^A-Za-z0-9 ._]", "", nome)

    # 6️⃣ Normaliza espaços e remove pontuação final
    nome = re.sub(r"\s+", " ", nome).strip()
    nome = nome.rstrip(" .")

    return nome or "Sem_Nome"

def salvar_captura_de_tela_declaracao(pagina, caminho, mes, ano):
    nome_arquivo = f"declaracao_sem_movimento_{str(mes).zfill(2)}.{ano}.png"
    caminho_arquivo = os.path.join(caminho, nome_arquivo)
    try:
        pagina.screenshot(path=caminho_arquivo, full_page=True)
        print(f"📸 Captura de tela salva em: {caminho_arquivo}")
    except Exception as e:
        print(f"❗ Erro ao salvar captura de tela: {e}")

def salvar_captura_de_tela(pagina, caminho, mes, ano, sufixo):
    nome_arquivo = f"sem_movimento_{str(mes).zfill(2)}.{ano}_{sufixo}.png"
    caminho_arquivo = os.path.join(caminho, nome_arquivo)
    try:
        pagina.screenshot(path=caminho_arquivo, full_page=True)
        print(f"📸 Captura de tela salva em: {caminho_arquivo}")
    except Exception as e:
        print(f"❗ Erro ao salvar captura de tela: {e}")

def salvar_captura_erro(pagina, caminho, mes, ano, sufixo):
    nome_arquivo = f"ERRO DECLARAÇÃO PREFEITURA {str(mes).zfill(2)}.{ano} - {sufixo}.png"
    caminho_arquivo = os.path.join(caminho, nome_arquivo)
    pagina.screenshot(path=caminho_arquivo, full_page=True)
    print(f"📸 Captura de ERRO salva em: {caminho_arquivo}")

def baixar_arquivos(pagina, nome_prestador, mes_extenso, ano_ref, mes_anterior, origem_texto, tem_registro, valor_prestador):
    time.sleep(1.5)
    pagina.click("text=Pesquisar")

    try:
        pagina.wait_for_load_state("domcontentloaded")
        if pagina.is_visible("text=Não há registros"):
            tem_registro = False
        elif pagina.locator("table#tabelaDinamica i.fa.fa-search").count() > 0:
            tem_registro = True
        else:
            tem_registro = False
    except Exception as e:
        print(f"⚠️ Erro ao verificar movimento: {e}")
        tem_registro = False

    sufixo = "emitido" if origem_texto.lower() == "emitida" else "recebido"

    nome_limpo = nome_prestador.split(" - ", 1)[1] if " - " in nome_prestador else nome_prestador
    nome_limpo = limpar_nome(nome_limpo)
    pasta_cliente = os.path.join(base_download_dir, nome_limpo)
    pasta_mes_ano = os.path.join(pasta_cliente, f"{str(mes_anterior).zfill(2)}.{ano_ref}")
    os.makedirs(pasta_mes_ano, exist_ok=True)

    try:
        if tem_registro:
            with pagina.expect_download(timeout=30000) as download_info:
                pagina.click("text=Exportar em XML")
            download = download_info.value
            nome_arquivo_xml = f"notas_{mes_extenso.lower()}_{ano_ref}_{sufixo}.xml"
        else:
            salvar_captura_de_tela(pagina, pasta_mes_ano, mes_anterior, ano_ref, sufixo)

            if sufixo == "emitido":
                try:
                    print(f"⚠️ Sem registros — emitindo declaração sem movimento.")
                    registrar_log(nome_prestador, f"SEM REGISTROS {sufixo.upper()} {str(mes_anterior).zfill(2)}/{ano_ref} — iniciando declaração sem movimento")

                    prestador_select = pagina.locator('select[name="parametrosTela.idPessoa"]')
                    prestador_select.select_option(value=valor_prestador)

                    pagina.wait_for_selector('select[name="parametrosTela.nrMesCompetencia"]')
                    pagina.select_option('select[name="parametrosTela.nrMesCompetencia"]', label=mes_extenso)
                    pagina.select_option('select[name="parametrosTela.nrAnoCompetencia"]', str(ano_ref))

                except Exception as e:
                    print(f"❌ Erro ao executar declaracao.py: {e}")

            nome_arquivo_xml = f"sem_movimento_{mes_extenso.lower()}_{ano_ref}_{sufixo}.xml"
            download = None

        caminho_final_xml = os.path.join(pasta_mes_ano, nome_arquivo_xml)
        if download:
            download.save_as(caminho_final_xml)
            print(f"✅ XML salvo em:\n{caminho_final_xml}")
        else:
            print(f"ℹ️ Nenhum XML gerado")
            registrar_log(nome_prestador, f"NENHUM XML GERADO {sufixo.upper()} {str(mes_anterior).zfill(2)}/{ano_ref}")

        return tem_registro, pasta_mes_ano

    except Exception as e:
        print(f"⚠️ Falha ao exportar XML: {e}")
        registrar_log(nome_prestador, f"ERRO XML {sufixo.upper()} {str(mes_anterior).zfill(2)}/{ano_ref}: {e}")
        return False, pasta_mes_ano

def baixar_relatorio(pagina, mes_extenso, ano_ref, mes_anterior, pasta_mes_ano,tem_registro_emitida, tem_registro_recebida, valor_prestador):
    try:
        if tem_registro_emitida:
            prestador_select = pagina.locator('select[name="formulario.idPessoa"]')
            prestador_select.select_option(value=valor_prestador)
            pagina.select_option('select[name="formulario.tpOrigemNfs"]', label="Emitida")
            pagina.select_option('select[name="formulario.nrMesCompetencia"]', label=str(mes_anterior))
            pagina.select_option('select[name="formulario.nrAnoCompetencia"]', str(ano_ref))
            pagina.locator('input[value="Pesquisar"]').click()

            pagina.wait_for_load_state("domcontentloaded")
            try:
                with pagina.expect_download(timeout=30000) as download_info:
                    pagina.locator('span.pdf.fa.fa-file-pdf-o').click()
                download_pdf = download_info.value
                nome_arquivo_pdf = f"notas_{mes_extenso.lower()}_{ano_ref}_emitido.pdf"
                caminho_final_pdf = os.path.join(pasta_mes_ano, nome_arquivo_pdf)
                download_pdf.save_as(caminho_final_pdf)
                print(f"✅ PDF EMITIDA salvo em:{caminho_final_pdf}")
                registrar_log("EMITIDA - " + valor_prestador, f"SUCESSO PDF EMITIDO {str(mes_anterior).zfill(2)}/{ano_ref} | {caminho_final_pdf}")
            except Exception as e:
                print(f"⚠️ Falha ao exportar PDF Emitida: {e}")
                registrar_log("EMITIDA - " + valor_prestador, f"ERRO PDF EMITIDO {str(mes_anterior).zfill(2)}/{ano_ref}: {e}")

            pagina.click("text=Limpar")
            pagina.wait_for_load_state("domcontentloaded")
    except Exception as e:
        print(f"⚠️ Erro ao processar relatórios 'Emitida': {e}")

    try:
        if tem_registro_recebida:
            try:
                pagina.wait_for_selector('select[name="formulario.idPessoa"]', timeout=30000)
                prestador_select = pagina.locator('select[name="formulario.idPessoa"]')
                prestador_select.select_option(value=valor_prestador)
            except Exception:
                raise RuntimeError("Select 'formulario.idPessoa' não encontrado na aba RECEBIDA; retornando para tela de consulta.")

            pagina.select_option('select[name="formulario.tpOrigemNfs"]', label="Recebida")
            pagina.select_option('select[name="formulario.nrMesCompetencia"]', label=str(mes_anterior))
            pagina.select_option('select[name="formulario.nrAnoCompetencia"]', str(ano_ref))
            pagina.locator('input[value="Pesquisar"]').click()

            pagina.wait_for_load_state("domcontentloaded")
            try:
                with pagina.expect_download(timeout=30000) as download_info:
                    pagina.locator('span.pdf.fa.fa-file-pdf-o').click()
                download_pdf = download_info.value
                nome_arquivo_pdf = f"notas_{mes_extenso.lower()}_{ano_ref}_recebido.pdf"
                caminho_final_pdf = os.path.join(pasta_mes_ano, nome_arquivo_pdf)
                download_pdf.save_as(caminho_final_pdf)
                print(f"✅ PDF RECEBIDA salvo em:{caminho_final_pdf}")
                registrar_log("RECEBIDA - " + valor_prestador, f"SUCESSO PDF RECEBIDO {str(mes_anterior).zfill(2)}/{ano_ref} | {caminho_final_pdf}")
            except Exception as e:
                print(f"⚠️ Falha ao exportar PDF Recebida: {e}")
                registrar_log("RECEBIDA - " + valor_prestador, f"ERRO PDF RECEBIDO {str(mes_anterior).zfill(2)}/{ano_ref}: {e}")

            pagina.click("text=Limpar")
            pagina.wait_for_load_state("domcontentloaded")
    except RuntimeError as e:
        print(f"ℹ️ {e}")
    except Exception as e:
        print(f"⚠️ Erro ao processar relatórios 'Recebida': {e}")

    try:
        time.sleep(1.5)
        pagina.click("text=NFS-E")
        pagina.click("text=Consulta")

        pagina.wait_for_load_state("domcontentloaded")

        pagina.wait_for_function(
            """() => {
                const select = document.querySelector('select[name="parametrosTela.idPessoa"]');
                return select && select.options.length > 1;
            }"""
        )

        pagina.wait_for_selector('select[name="parametrosTela.idPessoa"]')
        prestador_select = pagina.locator('select[name="parametrosTela.idPessoa"]')
        prestador_select.select_option(value=valor_prestador)

        pagina.wait_for_selector('select[name="parametrosTela.nrMesCompetencia"]')
        pagina.select_option('select[name="parametrosTela.nrMesCompetencia"]', label=mes_extenso)
        pagina.select_option('select[name="parametrosTela.nrAnoCompetencia"]', str(ano_ref))

        pagina.wait_for_load_state("domcontentloaded")
    except Exception as e:
        print(f"❗ Erro ao retornar para tela de consulta: {e}")


with sync_playwright() as p:
    contexto = p.firefox.launch_persistent_context(
        user_data_dir=r"C:\ROBOS\perfis firefox\BUSCA NOTAS\perfil_firefox_cert_esnfs",
        headless=False,
        accept_downloads=True,
        firefox_user_prefs={
            # Mantém a seleção automática do certificado salva no perfil
            "security.default_personal_cert": "Select Automatically",
            "security.remember_cert_checkbox_default_setting": True,
        },
    )
    pagina = contexto.new_page()

    pagina.goto("https://www.esnfs.com.br/?e=35")
    pagina.wait_for_load_state("domcontentloaded")

    time.sleep(2)

    # Login por CERTIFICADO DIGITAL (o perfil já lembra o certificado escolhido)
    pagina.wait_for_selector("text=Certificado digital", timeout=30000)
    pagina.click("text=Certificado digital")

    # Seleciona o município (como no fluxo original)
    pagina.wait_for_selector("text=Município de Francisco Beltrão", timeout=30000)
    pagina.click("text=Município de Francisco Beltrão")
    pagina.wait_for_load_state("domcontentloaded")

    # Dá alguns segundos para ver se realmente logou
    time.sleep(4)
    url_atual = pagina.url
    if "login" in url_atual.lower() or "captcha" in url_atual.lower() or pagina.locator("iframe[title*='recaptcha']").count() > 0:
        print("🚫 reCAPTCHA bloqueou o login — reiniciando em 3s...")
        contexto.close()
        time.sleep(3)
        subprocess.run([sys.executable, sys.argv[0]], check=True)
        sys.exit(0)

    print("✅ Login efetuado com sucesso, prosseguindo...")
    # ========= FIM LOGIN =========

    # Navegação para tela de consulta (após login)
    pagina.wait_for_load_state("domcontentloaded")
    pagina.click("text=NFS-E")
    pagina.click("text=Consulta")
    pagina.wait_for_load_state("domcontentloaded")

    pagina.wait_for_selector('select[name="parametrosTela.idPessoa"]')
    prestador_select = pagina.locator('select[name="parametrosTela.idPessoa"]')

    prestadores = prestador_select.locator("option").all()
    prestadores_info = [(p.get_attribute("value"), p.text_content().strip()) for p in prestadores if p.get_attribute("value")]

    # ========== BLOCO de escolha do início (já adicionado anteriormente) ==========
    def _normalizar(txt: str) -> str:
        return unicodedata.normalize("NFKD", txt or "").encode("ASCII","ignore").decode("ASCII").lower().strip()

    def _extrair_cnpj(texto_opcao: str) -> str | None:
        m = re.search(r'(\d{14})', re.sub(r'\D', '', texto_opcao))
        if m:
            return m.group(1)
        m2 = re.search(r'(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})', texto_opcao)
        if m2:
            return re.sub(r'\D', '', m2.group(1))
        return None

    indice_por_valor = {}
    indice_por_cnpj  = {}
    for idx, (valor, texto) in enumerate(prestadores_info):
        indice_por_valor[valor] = idx
        cnpj = _extrair_cnpj(texto)
        if cnpj:
            indice_por_cnpj[cnpj] = idx

    print("\n➤ Você pode informar:\n"
          "  • CNPJ (com ou sem máscara)\n"
          "  • Código 'value' (do select)\n"
          "  • Parte do nome da empresa\n"
          "  • Ou Enter para iniciar como antes (ID 0, se existir).")

    entrada = input("\nDigite aqui qual empresa deseja iniciar (ou Enter p/ manter padrão): ").strip()

    if entrada:
        entrada_digits = re.sub(r'\D', '', entrada)
        indice_inicio = None
        if entrada in indice_por_valor:
            indice_inicio = indice_por_valor[entrada]
            print(f"➡️ Iniciando pelo código (value) '{entrada}'.")
        if indice_inicio is None and len(entrada_digits) == 14 and entrada_digits in indice_por_cnpj:
            indice_inicio = indice_por_cnpj[entrada_digits]
            print(f"➡️ Iniciando pelo CNPJ {entrada_digits}.")
        if indice_inicio is None:
            alvo = _normalizar(entrada)
            for i, (_, texto) in enumerate(prestadores_info):
                if alvo and _normalizar(texto).find(alvo) != -1:
                    indice_inicio = i
                    print(f"➡️ Iniciando pelo nome que contém: '{entrada}'.")
                    break
        if indice_inicio is None:
            print("⚠️ Não encontrei essa empresa. Prosseguindo do jeito atual (ID 0, se houver).")
            indice_inicio = next((i for i, (valor, _) in enumerate(prestadores_info) if valor == '0'), 0)
    else:
        indice_inicio = next((i for i, (valor, _) in enumerate(prestadores_info) if valor == '0'), 0)
        print("↩️ Nenhuma escolha informada. Iniciando do jeito atual (a partir do ID 0, quando existir).")

    prestadores_a_processar = prestadores_info[indice_inicio:]
    if not prestadores_a_processar:
        raise Exception("❌ Nenhum prestador válido encontrado para processar a partir do índice escolhido.")

    hoje = datetime.today()
    mes_anterior = hoje.month - 1 or 12
    ano_ref = hoje.year if hoje.month > 1 else hoje.year - 1
    meses_ext = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mes_extenso = meses_ext[mes_anterior - 1]

    for valor_prestador, nome_prestador_completo in prestadores_a_processar:
        print(f"\n🔍 Processando prestador: {nome_prestador_completo} (valor: {valor_prestador})")

        time.sleep(1.5)
        pagina.click("text=Limpar")
        time.sleep(1.5)

        prestador_select = pagina.locator('select[name="parametrosTela.idPessoa"]')
        prestador_select.select_option(value=valor_prestador)

        pagina.wait_for_selector('select[name="parametrosTela.nrMesCompetencia"]')
        pagina.select_option('select[name="parametrosTela.nrMesCompetencia"]', label=mes_extenso)
        pagina.select_option('select[name="parametrosTela.nrAnoCompetencia"]', str(ano_ref))

        pagina.select_option('select[name="parametrosTela.origemEmissaoNfse"]', label="Emitida")
        pagina.wait_for_load_state("domcontentloaded")
        tem_registro_emitida, pasta_mes_ano = baixar_arquivos(
            pagina, nome_prestador_completo, mes_extenso, ano_ref, mes_anterior, "Emitida", True, valor_prestador
        )
        time.sleep(1.5)
        pagina.click("text=Limpar")
        time.sleep(1.5)

        pagina.select_option('select[name="parametrosTela.origemEmissaoNfse"]', label="Recebida")
        pagina.wait_for_selector('select[name="parametrosTela.nrMesCompetencia"]')
        pagina.select_option('select[name="parametrosTela.nrMesCompetencia"]', label=mes_extenso)
        pagina.select_option('select[name="parametrosTela.nrAnoCompetencia"]', str(ano_ref))
        pagina.wait_for_load_state("domcontentloaded")
        tem_registro_recebida, pasta_mes_ano = baixar_arquivos(
            pagina, nome_prestador_completo, mes_extenso, ano_ref, mes_anterior, "Recebida", True, valor_prestador
        )

        if tem_registro_emitida is True or tem_registro_recebida is True:
            time.sleep(1.5)
            pagina.click("text=RELATÓRIOS")
            pagina.click("text=Apuração do ISS")
            time.sleep(3)   

            try:
                prestador_select = pagina.locator('select[name="formulario.idPessoa"]')
                prestador_select.select_option(value=valor_prestador)

                pagina.wait_for_selector('select[name="formulario.nrMesCompetencia"]')
                pagina.select_option('select[name="formulario.nrMesCompetencia"]', label=str(mes_anterior))
                pagina.select_option('select[name="formulario.nrAnoCompetencia"]', str(ano_ref))

                baixar_relatorio(
                    pagina,
                    mes_extenso,
                    ano_ref,
                    mes_anterior,
                    pasta_mes_ano,
                    tem_registro_emitida,
                    tem_registro_recebida,
                    valor_prestador
                )

            except Exception as e:
                print(f"❗ Prestador com valor '{valor_prestador}' não encontrado nas opções do relatório. Pulando esta etapa.")
                print(f"Detalhes do erro: {e}")

            try:
                time.sleep(1.5)
                pagina.click("text=NFS-E")
                pagina.click("text=Consulta")
                time.sleep(3)
            except Exception as e:
                print(f"❗ Erro ao retornar para tela de consulta após o relatório: {e}")

    CAMINHO_PRINCIPAL = r"C:\Users\Usuario\Documents\PYTHON\IMPORTADOR_NFSE\principal.py"
    print("\n🚀 Executando principal.py...")
    subprocess.run(["python", CAMINHO_PRINCIPAL], check=True)
315