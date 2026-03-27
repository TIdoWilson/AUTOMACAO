import os
import csv
import time
import unicodedata
import re
from datetime import datetime
from playwright.sync_api import sync_playwright
import pyautogui

# Diretório base dos downloads
BASE_DOWNLOAD_DIR = r"\\192.0.0.251\arquivos\XML PREFEITURA"

# Perfil persistente do Firefox (mesmo usado no buscador de notas)
FIREFOX_PROFILE_DIR = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\BUSCADOR DE NOTAS\perfil_firefox_cert_esnfs"

# Lista de log dos prestadores
log_prestadores = []

# Utilitário: monta o caminho final do PDF
def limpar_nome(nome: str) -> str:
    nome = unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("ASCII")
    nome = re.sub(r"[^A-Za-z0-9 .]", "", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome.rstrip(" .")

def montar_caminho_download(nome_prestador, mes, ano):
    nome_limpo = nome_prestador.split(" - ", 1)[1] if " - " in nome_prestador else nome_prestador
    nome_limpo = limpar_nome(nome_limpo)   # <<< usar a mesma limpeza
    pasta_cliente = os.path.join(BASE_DOWNLOAD_DIR, nome_limpo)
    pasta_mes_ano = os.path.join(pasta_cliente, f"{str(mes).zfill(2)}.{ano}")
    os.makedirs(pasta_mes_ano, exist_ok=True)
    nome_arquivo_pdf = f"ISS_{str(mes).zfill(2)}.{ano}.pdf"   # padronizado
    caminho_final_pdf = os.path.join(pasta_mes_ano, nome_arquivo_pdf)
    return caminho_final_pdf


# Salva log CSV
def salvar_log_em_csv():
    caminho_csv = os.path.join(BASE_DOWNLOAD_DIR, "log_emissao_guias.csv")
    with open(caminho_csv, mode="w", newline="", encoding="utf-8") as arquivo_csv:
        campos = ["Prestador", "Pesquisa", "Clique Emitir", "Download Guia", "Mensagem de Erro"]
        writer = csv.DictWriter(arquivo_csv, fieldnames=campos)
        writer.writeheader()
        for linha in log_prestadores:
            writer.writerow(linha)
    print(f"\n📝 Log salvo em: {caminho_csv}")

# Emite todas as guias disponíveis para um prestador
def emitir_guias(pagina, contexto, nome_prestador, mes, ano):
    registro = {
        "Prestador": nome_prestador,
        "Pesquisa": "OK",
        "Clique Emitir": "-",
        "Download Guia": "-",
        "Mensagem de Erro": ""
    }

    links_emissao = pagina.locator('a[title="Emissão"]')
    total_links = links_emissao.count()

    if total_links == 0:
        print("⚠️ Nenhuma guia para emitir.")
        log_prestadores.append(registro)
        return

    # Filtra pelos do mês/ano correto
    hrefs_filtrados = []
    for i in range(total_links):
        link = links_emissao.nth(i)
        href = link.get_attribute("href")
        if href and "viewEditGuia" in href:
            partes = href.split("'")
            if len(partes) >= 6:
                mes_href = int(partes[3])
                ano_href = int(partes[5])
                if mes_href == mes and ano_href == ano:
                    hrefs_filtrados.append(link)

    if not hrefs_filtrados:
        print(f"⚠️ Nenhuma guia encontrada para {mes}/{ano}")
        registro["Mensagem de Erro"] = f"Guia(s) de {mes}/{ano} não encontrada(s)"
        log_prestadores.append(registro)
        return

    while True:
        try:
            link = hrefs_filtrados[0]
            href = link.get_attribute("href")

            if href and "viewEditGuia" in href:
                link.locator("i.fa.fa-barcode").click()
                print("🔘 Clicando no botão de emissão...")
                registro["Clique Emitir"] = "OK"

                pagina.wait_for_selector("input#emitir", timeout=5000)
                pagina.locator("input#emitir").scroll_into_view_if_needed()
                pagina.wait_for_timeout(300)

                pagina.once("dialog", lambda dialog: dialog.accept())

                try:
                    with contexto.expect_page(timeout=5000) as nova_guia_info:
                        pagina.locator("input#emitir").click()

                    nova_pagina = nova_guia_info.value
                    nova_pagina.wait_for_url("**/nfsguiarecolhimento.imprimir.logic", timeout=10000)
                    nova_pagina.wait_for_load_state("load")
                    print("📄 PDF carregado. Tentando fazer download...")
                    time.sleep(3)

                    with nova_pagina.expect_download(timeout=10000) as download_info:
                        # Ajuste as coordenadas conforme seu viewer de PDF
                        pyautogui.moveTo(1192, 124, duration=0.1)
                        pyautogui.click()
                        time.sleep(2)

                    download_pdf = download_info.value
                    caminho_final_pdf = montar_caminho_download(nome_prestador, mes, ano)
                    download_pdf.save_as(caminho_final_pdf)
                    print(f"✅ PDF salvo em:\n{caminho_final_pdf}")
                    registro["Download Guia"] = "OK"

                    nova_pagina.close()
                    pagina.bring_to_front()
                    
                    time.sleep(3)

                    try:
                        pyautogui.moveTo(880, 225, duration=0.1)
                        pyautogui.click()
                        time.sleep(2)
                        
                        pagina.click('input.botaoVoltar')
                        pagina.wait_for_load_state("networkidle")
                        pagina.wait_for_selector('select[name="pessoaModel.idPessoa"]:not([disabled])', timeout=5000)
                        time.sleep(1)
                    except Exception as e:
                        registro["Mensagem de Erro"] = f"Erro ao retornar: {e}"
                        break

                    if pagina.is_visible("text=Não há registros"):
                        print("✅ Nenhuma nova guia encontrada. Passando para o próximo prestador.")
                        break

                except Exception as e:
                    msg = "Guia já emitida ou erro ao abrir nova aba"
                    print(f"⚠️ {msg}: {e}")
                    registro["Mensagem de Erro"] = msg

                    try:
                        pyautogui.moveTo(880, 225, duration=0.1)
                        pyautogui.click()
                        pagina.click('input.botaoVoltar')
                        pagina.wait_for_load_state("networkidle")
                        time.sleep(1)
                    except Exception as e2:
                        print(f"❌ Falha ao tentar voltar após erro: {e2}")
                        registro["Mensagem de Erro"] += f" | Falha ao voltar: {e2}"
                    break

        except Exception as e:
            print(f"❗ Erro inesperado ao emitir guia: {e}")
            registro["Mensagem de Erro"] = str(e)
            break

    log_prestadores.append(registro)

# Processa todos os prestadores
def processar_prestadores(pagina, contexto):
    hoje = datetime.today()
    mes_anterior = hoje.month - 1 if hoje.month > 1 else 12
    ano_ref = hoje.year if hoje.month > 1 else hoje.year - 1

    prestador_select = pagina.locator('select[name="pessoaModel.idPessoa"]')
    prestadores = prestador_select.locator("option").all()
    total_prestadores = len(prestadores)

    if total_prestadores < 2:
        raise Exception("❌ Nenhum prestador válido encontrado.")

    for index in range(1, total_prestadores):
        nome_prestador = prestadores[index].text_content().strip()
        print(f"\n🔍 Processando prestador: {nome_prestador}")

        registro = {
            "Prestador": nome_prestador,
            "Pesquisa": "SEM DADOS",
            "Clique Emitir": "-",
            "Download Guia": "-",
            "Mensagem de Erro": ""
        }

        prestador_select.select_option(index=index)
        pagina.wait_for_selector('select[name="formulario.nrExercicio"]')
        pagina.select_option('select[name="formulario.nrExercicio"]', str(ano_ref))
        pagina.click("text=Pesquisar")
        pagina.wait_for_timeout(1500)

        if pagina.is_visible("text=Não há registros"):
            print("❌ Nenhum registro encontrado.")
            registro["Pesquisa"] = "SEM REGISTROS"
            log_prestadores.append(registro)
            continue

        registro["Pesquisa"] = "OK"
        try:
            links_emissao = pagina.locator('a[title="Emissão"]')
            total_links = links_emissao.count()

            if total_links > 0:
                print(f"✅ {total_links} guia(s) localizada(s)")
                emitir_guias(pagina, contexto, nome_prestador, mes_anterior, ano_ref)
            else:
                print("⚠️ Nenhum link de emissão encontrado.")
                registro["Mensagem de Erro"] = "Sem links de emissão"
                log_prestadores.append(registro)
        except Exception as e:
            registro["Mensagem de Erro"] = str(e)
            log_prestadores.append(registro)

# Fluxo principal
def main():
# =================== INICIALIZAÇÃO (perfil Firefox + certificado) ===================
    with sync_playwright() as p:
        contexto = p.firefox.launch_persistent_context(
            user_data_dir=r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\BUSCADOR DE NOTAS\perfil_firefox_cert_esnfs",
            headless=False,
            accept_downloads=True,
            firefox_user_prefs={
                "security.default_personal_cert": "Select Automatically",
                "security.remember_cert_checkbox_default_setting": True,
            },
        )
        pagina = contexto.new_page()

        pagina.goto("https://www.esnfs.com.br/?e=35")
        pagina.wait_for_load_state("domcontentloaded")

        time.sleep(2)

        # Fecha modal "Fechar" se aparecer
        try:
            pagina.locator("div.modal-footer button[data-dismiss='modal'], button:has-text('Fechar')").first.click()
            pagina.wait_for_load_state("domcontentloaded")
        except Exception:
            pass

        # Login por CERTIFICADO DIGITAL (o perfil já lembra o certificado escolhido)
        pagina.wait_for_selector("text=Certificado digital", timeout=30000)
        pagina.click("text=Certificado digital")

        # Seleciona o município
        pagina.wait_for_selector("text=Município de Francisco Beltrão", timeout=30000)
        pagina.click("text=Município de Francisco Beltrão")
        pagina.wait_for_load_state("domcontentloaded")

        # Caminho até a tela de emissão/consulta
        pagina.click("text=GUIA DE RECOLHIMENTO")
        pagina.click("text=ISS devido / Consulta / Cancelamento")
        time.sleep(1)

        processar_prestadores(pagina, contexto)
        salvar_log_em_csv()
        input("\n🛑 Pressione ENTER para encerrar manualmente...")

if __name__ == "__main__":
    main()
