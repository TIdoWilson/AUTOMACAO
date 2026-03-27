import os
import csv
import time
import unicodedata
import re
from datetime import datetime
from playwright.sync_api import sync_playwright

BASE_DOWNLOAD_DIR = r"\\192.0.0.251\arquivos\XML PREFEITURA"
FIREFOX_PROFILE_DIR = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\BUSCADOR DE NOTAS\perfil_firefox_cert_esnfs"

ESNFS_ORIGIN = "https://www.esnfs.com.br"
URL_TELA_CONSULTA = f"{ESNFS_ORIGIN}/nfsguiarecolhimento.load.logic"

START_INDEX = 100
log_prestadores = []


def limpar_nome(nome: str) -> str:
    nome = unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("ASCII")
    nome = re.sub(r"[^A-Za-z0-9 .]", "", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome.rstrip(" .")


def montar_caminho_download(nome_prestador, mes, ano):
    nome_limpo = nome_prestador.split(" - ", 1)[1] if " - " in nome_prestador else nome_prestador
    nome_limpo = limpar_nome(nome_limpo)
    pasta_cliente = os.path.join(BASE_DOWNLOAD_DIR, nome_limpo)
    pasta_mes_ano = os.path.join(pasta_cliente, f"{str(mes).zfill(2)}.{ano}")
    os.makedirs(pasta_mes_ano, exist_ok=True)
    nome_arquivo_pdf = f"ISS_{str(mes).zfill(2)}.{ano}.pdf"
    return os.path.join(pasta_mes_ano, nome_arquivo_pdf)


def salvar_log_em_csv():
    caminho_csv = os.path.join(BASE_DOWNLOAD_DIR, "log_emissao_guias.csv")
    with open(caminho_csv, mode="w", newline="", encoding="utf-8") as arquivo_csv:
        campos = ["Prestador", "Pesquisa", "Clique Emitir", "Download Guia", "Mensagem de Erro"]
        writer = csv.DictWriter(arquivo_csv, fieldnames=campos)
        writer.writeheader()
        for linha in log_prestadores:
            writer.writerow(linha)
    print(f"\n📝 Log salvo em: {caminho_csv}")


def garantir_tela_consulta(contexto, pagina):
    if pagina is None or pagina.is_closed():
        pagina = contexto.new_page()

    try:
        pagina.keyboard.press("Escape")
        pagina.wait_for_timeout(150)
        pagina.keyboard.press("Escape")
    except Exception:
        pass

    pagina.goto(URL_TELA_CONSULTA, wait_until="domcontentloaded")
    pagina.wait_for_selector('select[name="pessoaModel.idPessoa"]', timeout=20000)
    return pagina


def voltar_para_consulta(contexto, pagina):
    try:
        if pagina is not None and not pagina.is_closed():
            if pagina.locator('input.botaoVoltar').count() > 0:
                pagina.click('input.botaoVoltar', timeout=8000)
                pagina.wait_for_load_state("networkidle")
    except Exception:
        pass

    return garantir_tela_consulta(contexto, pagina)


def baixar_pdf_no_viewer(popup, contexto, caminho_final_pdf: str):
    """
    Clica no botão "Download" do viewer do Firefox (PDF.js) e captura o download.
    """
    # Viewer costuma ter #download. Se não tiver, tenta variações.
    seletores = [
        "#download",                          # comum no PDF.js
        "button#download",
        "button[title='Download']",
        "button[title*='Download' i]",
        "text=Download",
        "text=Baixar",
        "button:has-text('Download')",
        "button:has-text('Baixar')",
        "a:has-text('Download')",
        "a:has-text('Baixar')",
    ]

    # garante que o viewer carregou
    popup.wait_for_load_state("domcontentloaded", timeout=15000)
    # às vezes demora para montar a toolbar
    popup.wait_for_timeout(800)

    last_err = None
    for sel in seletores:
        try:
            loc = popup.locator(sel)
            if loc.count() == 0:
                continue

            btn = loc.first
            btn.scroll_into_view_if_needed()

            # download pode ser gerado pelo popup; melhor capturar no CONTEXTO
            with contexto.expect_download(timeout=30000) as dlinfo:
                btn.click(force=True)

            dl = dlinfo.value
            dl.save_as(caminho_final_pdf)
            return True
        except Exception as e:
            last_err = e
            continue

    raise RuntimeError(f"Não consegui clicar no botão de download do viewer. Último erro: {last_err}")


def emitir_guias(contexto, pagina, nome_prestador, mes, ano):
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
        return pagina

    # Filtra pelo mês/ano no href
    candidatos = []
    for i in range(total_links):
        link = links_emissao.nth(i)
        href = link.get_attribute("href") or ""
        if "viewEditGuia" in href:
            partes = href.split("'")
            if len(partes) >= 6:
                try:
                    mes_href = int(partes[3])
                    ano_href = int(partes[5])
                except Exception:
                    continue
                if mes_href == mes and ano_href == ano:
                    candidatos.append(link)

    if not candidatos:
        print(f"⚠️ Nenhuma guia encontrada para {mes}/{ano}")
        registro["Mensagem de Erro"] = f"Guia(s) de {mes}/{ano} não encontrada(s)"
        log_prestadores.append(registro)
        return pagina

    link_alvo = candidatos[0]

    try:
        link_alvo.locator("i.fa.fa-barcode").click()
        print("🔘 Clicando no botão de emissão...")
        registro["Clique Emitir"] = "OK"

        pagina.wait_for_selector("input#emitir", timeout=15000)
        btn_emitir = pagina.locator("input#emitir")
        btn_emitir.scroll_into_view_if_needed()
        pagina.wait_for_timeout(250)

        # popup "OK/Cancelar"
        pagina.once("dialog", lambda dialog: dialog.accept())

        # ✅ Impede o portal de fechar a aba automaticamente
        # (aplica para páginas novas)
        try:
            contexto.add_init_script("window.close = () => {};")
        except Exception:
            pass

        # Clique REAL no Emitir, capturando o popup
        popup = None
        with pagina.expect_popup(timeout=15000) as pop:
            btn_emitir.click()
        popup = pop.value

        # se o portal tentar fechar via JS, agora não fecha
        try:
            popup.on("close", lambda: print("⚠️ A aba de impressão fechou (mesmo após bloquear window.close)."))
        except Exception:
            pass

        print("⬇️ Baixando pelo botão Download do viewer...")
        caminho_final_pdf = montar_caminho_download(nome_prestador, mes, ano)
        baixar_pdf_no_viewer(popup, contexto, caminho_final_pdf)

        print(f"✅ PDF salvo em:\n{caminho_final_pdf}")
        registro["Download Guia"] = "OK"

        try:
            if not popup.is_closed():
                popup.close()
        except Exception:
            pass

    except Exception as e:
        print(f"❗ Erro ao emitir/baixar guia: {e}")
        registro["Mensagem de Erro"] = str(e)

    # volta sempre
    try:
        pagina = voltar_para_consulta(contexto, pagina)
    except Exception as e:
        registro["Mensagem de Erro"] += f" | Falha ao voltar: {e}"

    log_prestadores.append(registro)
    return pagina


def processar_prestadores(contexto, pagina, start_index=1):
    hoje = datetime.today()
    mes_anterior = hoje.month - 1 if hoje.month > 1 else 12
    ano_ref = hoje.year if hoje.month > 1 else hoje.year - 1

    pagina = garantir_tela_consulta(contexto, pagina)
    total_prestadores = pagina.locator('select[name="pessoaModel.idPessoa"] option').count()

    if total_prestadores < 2:
        raise Exception("❌ Nenhum prestador válido encontrado.")

    for index in range(start_index, total_prestadores):
        try:
            pagina = garantir_tela_consulta(contexto, pagina)

            opt = pagina.locator('select[name="pessoaModel.idPessoa"] option').nth(index)
            nome_prestador = (opt.text_content() or "").strip()
            print(f"\n🔍 Processando prestador: {nome_prestador}")

            registro = {
                "Prestador": nome_prestador,
                "Pesquisa": "SEM DADOS",
                "Clique Emitir": "-",
                "Download Guia": "-",
                "Mensagem de Erro": ""
            }

            pagina.locator('select[name="pessoaModel.idPessoa"]').select_option(index=index)
            pagina.wait_for_selector('select[name="formulario.nrExercicio"]', timeout=10000)
            pagina.select_option('select[name="formulario.nrExercicio"]', str(ano_ref))
            pagina.click("text=Pesquisar")
            pagina.wait_for_timeout(1500)

            if pagina.is_visible("text=Não há registros"):
                print("❌ Nenhum registro encontrado.")
                registro["Pesquisa"] = "SEM REGISTROS"
                log_prestadores.append(registro)
                continue

            registro["Pesquisa"] = "OK"

            links_emissao = pagina.locator('a[title="Emissão"]')
            total_links = links_emissao.count()

            if total_links > 0:
                print(f"✅ {total_links} guia(s) localizada(s)")
                pagina = emitir_guias(contexto, pagina, nome_prestador, mes_anterior, ano_ref)
            else:
                print("⚠️ Nenhum link de emissão encontrado.")
                registro["Mensagem de Erro"] = "Sem links de emissão"
                log_prestadores.append(registro)

        except Exception as e:
            print(f"❗ Falha ao processar prestador no índice {index}: {e}")
            log_prestadores.append({
                "Prestador": f"(index={index})",
                "Pesquisa": "ERRO",
                "Clique Emitir": "-",
                "Download Guia": "-",
                "Mensagem de Erro": str(e)
            })
            try:
                pagina = garantir_tela_consulta(contexto, pagina)
            except Exception:
                pass
            continue

    return pagina


def main():
    with sync_playwright() as p:
        contexto = p.firefox.launch_persistent_context(
            user_data_dir=FIREFOX_PROFILE_DIR,
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

        try:
            pagina.locator("div.modal-footer button[data-dismiss='modal'], button:has-text('Fechar')").first.click()
            pagina.wait_for_load_state("domcontentloaded")
        except Exception:
            pass

        pagina.wait_for_selector("text=Certificado digital", timeout=30000)
        pagina.click("text=Certificado digital")

        pagina.wait_for_selector("text=Município de Francisco Beltrão", timeout=30000)
        pagina.click("text=Município de Francisco Beltrão")
        pagina.wait_for_load_state("domcontentloaded")

        pagina = garantir_tela_consulta(contexto, pagina)

        pagina = processar_prestadores(contexto, pagina, start_index=START_INDEX)
        salvar_log_em_csv()

        input("\n🛑 Pressione ENTER para encerrar manualmente...")


if __name__ == "__main__":
    main()
