import os
import re
import time
import argparse
import unicodedata
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException,
    NoSuchElementException,
)
from webdriver_manager.firefox import GeckoDriverManager


# ================== CONFIGURAÇÕES GERAIS ==================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
LIMPOS_DIR = os.path.join(DOWNLOAD_DIR, "arquivos limpos")

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(LIMPOS_DIR, exist_ok=True)

URL = "https://www3.mte.gov.br/sistemas/mediador/ConsultarInstColetivo"

DT_INICIO_PADRAO = "01/01/2025"
DT_FIM_PADRAO = "31/12/2025"

TIMEOUT_PAGINA = 45
TIMEOUT_RESULTADOS = 60


def normalizar_texto(txt: str) -> str:
    if not txt:
        return ""
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    return txt.lower().strip()


def codigo_para_nome_arquivo(codigo: str) -> str:
    return re.sub(r"[^\w\-]", "_", str(codigo))


# ================== DRIVER FIREFOX ==================

def get_firefox_driver(download_path: str):
    options = FirefoxOptions()
    # options.add_argument("-headless")  # se quiser rodar sem abrir janela

    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.dir", download_path)
    profile.set_preference("browser.download.useDownloadDir", True)
    profile.set_preference("browser.download.manager.showWhenStarting", False)

    profile.set_preference(
        "browser.helperApps.neverAsk.saveToDisk",
        "application/pdf,application/msword,"
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    profile.set_preference("pdfjs.disabled", True)

    options.profile = profile

    driver = webdriver.Firefox(
        service=FirefoxService(GeckoDriverManager().install()),
        options=options
    )

    try:
        driver.maximize_window()
    except Exception:
        pass

    return driver


# ================== DOWNLOAD E RENAME ==================

def esperar_novo_arquivo(download_path: str, arquivos_antes, timeout: int = 60):
    limite = time.time() + timeout
    while time.time() < limite:
        arquivos_agora = set(os.listdir(download_path))
        novos = arquivos_agora - arquivos_antes

        candidatos = []
        for nome in novos:
            if nome.lower().endswith(".crdownload") or nome.lower().endswith(".part"):
                continue
            caminho = os.path.join(download_path, nome)
            if os.path.isfile(caminho):
                candidatos.append(caminho)

        if candidatos:
            return max(candidatos, key=os.path.getmtime)

        time.sleep(1)

    return None


def baixar_e_renomear(driver, link_element, download_path: str, codigo: str):
    arquivos_antes = set(os.listdir(download_path))
    main_window = driver.current_window_handle

    try:
        try:
            link_element.click()
        except Exception:
            driver.execute_script("arguments[0].click();", link_element)

        time.sleep(2)

        caminho_novo = esperar_novo_arquivo(download_path, arquivos_antes, timeout=60)

        if not caminho_novo or not os.path.exists(caminho_novo):
            print("      Erro: novo arquivo nao encontrado.")
            return None

        _, ext = os.path.splitext(caminho_novo)
        if not ext or ext.lower() in (".tmp", ".part"):
            ext = ".doc"

        codigo_limpo = codigo_para_nome_arquivo(codigo)
        caminho_final = os.path.join(download_path, f"{codigo_limpo}{ext}")

        if os.path.abspath(caminho_novo) == os.path.abspath(caminho_final):
            return caminho_final

        if os.path.exists(caminho_final):
            print(f"      Arquivo {os.path.basename(caminho_final)} ja existe.")
            return caminho_final

        try:
            os.rename(caminho_novo, caminho_final)
        except FileNotFoundError:
            print("      Erro ao renomear: arquivo sumiu.")
            return None

        return caminho_final

    finally:
        try:
            handles_depois = set(driver.window_handles)
            abas_extras = handles_depois - {main_window}
            for h in abas_extras:
                try:
                    driver.switch_to.window(h)
                    driver.close()
                except Exception:
                    pass
            driver.switch_to.window(main_window)
        except Exception:
            pass


# ================== EXTRACAO DOS DADOS DO BLOCO ==================

def extrair_dados_do_bloco(trs_bloco, link_element):
    codigo = ""
    codigo_onclick = ""
    sindicatos = ""
    data_inicio = ""
    data_fim = ""

    try:
        onclick = link_element.get_attribute("onclick") or ""
        m_on = re.search(r"fDownload\('([^']+)'", onclick)
        if m_on:
            codigo_onclick = m_on.group(1).strip()
    except Exception:
        onclick = ""

    textos_sind = []
    for tr in trs_bloco:
        try:
            tds = tr.find_elements(By.XPATH, ".//td[contains(@class,'textoConsulta2')]")
        except StaleElementReferenceException:
            raise
        except Exception:
            tds = []
        for td in tds:
            try:
                raw = td.get_attribute("innerText") or td.text
            except StaleElementReferenceException:
                raise
            except Exception:
                raw = ""
            if raw:
                txt = " ".join(raw.split())
                txt = re.sub(r"\s+e Outros$", "", txt, flags=re.IGNORECASE)
                if txt:
                    textos_sind.append(txt)

    if textos_sind:
        vistos = []
        for t in textos_sind:
            if t not in vistos:
                vistos.append(t)
        sindicatos = " | ".join(vistos)

    bloco_texto = ""
    for tr in trs_bloco:
        try:
            txt = tr.text
        except StaleElementReferenceException:
            raise
        except Exception:
            txt = ""
        if txt:
            bloco_texto += "\n" + txt

    bloco_texto_norm = " ".join(
        bloco_texto.replace("–", "-").replace("−", "-").split()
    )

    if bloco_texto_norm:
        m2 = re.search(r"(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})", bloco_texto_norm)
        if m2:
            data_inicio, data_fim = m2.group(1), m2.group(2)
        else:
            datas = re.findall(r"\d{2}/\d{2}/\d{4}", bloco_texto_norm)
            if len(datas) > 0:
                data_inicio = datas[0]
            if len(datas) > 1:
                data_fim = datas[1]

    codigo_registro = ""
    if bloco_texto_norm:
        m_reg = re.search(
            r"Registro\s+([A-Z]{2}\d{4,}/\d{4})",
            bloco_texto_norm,
            flags=re.IGNORECASE,
        )
        if m_reg:
            codigo_registro = m_reg.group(1).strip()
        else:
            all_codes = re.findall(r"\b[A-Z]{2}\d{4,}/\d{4}\b", bloco_texto_norm)
            for c in all_codes:
                if not c.startswith("MR"):
                    codigo_registro = c
                    break

    tr_principal = trs_bloco[0] if trs_bloco else None
    if tr_principal is not None:
        try:
            row_text = tr_principal.text.strip()
        except StaleElementReferenceException:
            raise
        except Exception:
            row_text = ""
    else:
        row_text = ""

    if row_text:
        if not codigo_registro:
            row_norm = " ".join(row_text.split())
            m_reg2 = re.search(
                r"Registro\s+([A-Z]{2}\d{4,}/\d{4})",
                row_norm,
                flags=re.IGNORECASE,
            )
            if m_reg2:
                codigo_registro = m_reg2.group(1).strip()
            else:
                all_codes2 = re.findall(r"\b[A-Z]{2}\d{4,}/\d{4}\b", row_norm)
                for c in all_codes2:
                    if not c.startswith("MR"):
                        codigo_registro = c
                        break

        if not data_inicio or not data_fim:
            datas2 = re.findall(r"\d{2}/\d{2}/\d{4}", row_text)
            if not data_inicio and len(datas2) > 0:
                data_inicio = datas2[0]
            if not data_fim and len(datas2) > 1:
                data_fim = datas2[1]

        if not sindicatos:
            linhas_texto = [t.strip() for t in row_text.splitlines() if t.strip()]
            partes_sind = []
            for t in linhas_texto:
                if codigo_registro and codigo_registro in t:
                    continue
                if "Download" in t:
                    continue
                if re.fullmatch(r"\d{2}/\d{2}/\d{4}", t):
                    continue
                partes_sind.append(t)
            if partes_sind:
                sindicatos = " | ".join(partes_sind)

    if codigo_registro:
        codigo = codigo_registro
    elif codigo_onclick:
        codigo = codigo_onclick
    else:
        codigo = ""

    return codigo, sindicatos, data_inicio, data_fim


# ================== REGISTRO NO EXCEL ==================

def registrar_download(df_registros, codigo, sindicatos, data_inicio, data_fim):
    novo = {
        "codigo": str(codigo),
        "sindicatos": str(sindicatos),
        "data_inicio": str(data_inicio),
        "data_fim": str(data_fim),
    }
    df_registros.loc[len(df_registros)] = novo
    df_registros.drop_duplicates(subset=["codigo"], keep="first", inplace=True)
    return df_registros


# ================== PAGINACAO ==================

def obter_paginacao_h2(driver):
    try:
        h2 = driver.find_element(
            By.XPATH,
            "//h2[contains(., 'Instrumento(s) Coletivo(s) Encontrado(s)')]"
        )
    except NoSuchElementException:
        return None, None
    except Exception:
        return None, None

    texto = h2.text or ""
    m = re.search(r"Pagina\s+(\d+)\s+de\s+(\d+)", texto)
    if not m:
        return None, None

    try:
        atual = int(m.group(1))
        total = int(m.group(2))
        return atual, total
    except ValueError:
        return None, None


def esta_processando(driver) -> bool:
    try:
        imagens = driver.find_elements(
            By.XPATH,
            "//img[contains(@src, 'imgProcessando') "
            "or contains(translate(@alt,'CARREGANDO','carregando'),'carregando')]"
        )
        for img in imagens:
            try:
                if img.is_displayed():
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def tentar_ir_para_proxima_pagina(driver, codigos_pagina_atual):
    pag_atual, total_paginas = obter_paginacao_h2(driver)

    if pag_atual is not None and total_paginas is not None:
        if pag_atual >= total_paginas:
            return False

    codigos_atual_chave = sorted(set(codigos_pagina_atual))

    try:
        link_proximo = driver.find_element(
            By.XPATH,
            "//a[contains(normalize-space(.),'Proximo') or contains(normalize-space(.),'Próximo')]"
        )
    except NoSuchElementException:
        return False
    except Exception:
        return False

    if not (link_proximo.is_displayed() and link_proximo.is_enabled()):
        return False

    try:
        driver.execute_script("arguments[0].click();", link_proximo)
    except Exception:
        try:
            link_proximo.click()
        except Exception:
            return False

    limite = time.time() + TIMEOUT_PAGINA
    while time.time() < limite:
        try:
            if esta_processando(driver):
                time.sleep(1)
                continue

            novo_pag, novo_total = obter_paginacao_h2(driver)
            if pag_atual is not None and novo_pag is not None and novo_pag != pag_atual:
                return True

            # como segurança, confere mudança nos codigos da tabela:
            try:
                tabela = driver.find_element(By.XPATH, "//table[.//a[contains(@onclick,'fDownload')]]")
                links = tabela.find_elements(By.XPATH, ".//a[contains(@onclick,'fDownload')]")
                codigos_novos = []
                for lk in links:
                    onclick = lk.get_attribute("onclick") or ""
                    m = re.search(r"fDownload\('([^']+)'", onclick)
                    if m:
                        codigos_novos.append(m.group(1).strip())
                if sorted(set(codigos_novos)) != codigos_atual_chave and codigos_novos:
                    return True
            except Exception:
                pass

        except StaleElementReferenceException:
            time.sleep(1)
            continue
        except Exception:
            time.sleep(1)
            continue

        time.sleep(1)

    return False


# ================== PROCESSAR UMA PAGINA ==================

def processar_pagina_atual(driver, df_registros, codigos_baixados):
    time.sleep(1)

    try:
        tabela = driver.find_element(By.XPATH, "//table[.//a[contains(@onclick,'fDownload')]]")
    except NoSuchElementException:
        print("    Nenhuma tabela de resultados encontrada.")
        return [], df_registros, codigos_baixados
    except Exception:
        print("    Erro ao localizar tabela de resultados.")
        return [], df_registros, codigos_baixados

    trs = tabela.find_elements(By.XPATH, ".//tr")
    i = 0
    codigos_pagina = []
    codigos_ja_processados_na_pagina = set()
    novos_registros_pagina = 0

    while i < len(trs):
        tr = trs[i]
        try:
            links = tr.find_elements(By.XPATH, ".//a[contains(@onclick,'fDownload')]")
        except StaleElementReferenceException:
            raise
        except Exception:
            links = []

        if not links:
            i += 1
            continue

        link_download = links[0]

        codigo_tmp = ""
        try:
            onclick_tmp = link_download.get_attribute("onclick") or ""
            m_tmp = re.search(r"fDownload\('([^']+)'", onclick_tmp)
            if m_tmp:
                codigo_tmp = m_tmp.group(1).strip()
        except Exception:
            codigo_tmp = ""

        if codigo_tmp and codigo_tmp in codigos_ja_processados_na_pagina:
            j = i + 1
            while j < len(trs):
                tr_next = trs[j]
                try:
                    links_next = tr_next.find_elements(By.XPATH, ".//a[contains(@onclick,'fDownload')]")
                except StaleElementReferenceException:
                    raise
                except Exception:
                    links_next = []
                if links_next:
                    break
                j += 1
            i = j
            continue

        bloco = [tr]
        j = i + 1
        while j < len(trs):
            tr_next = trs[j]
            try:
                links_next = tr_next.find_elements(By.XPATH, ".//a[contains(@onclick,'fDownload')]")
            except StaleElementReferenceException:
                raise
            except Exception:
                links_next = []
            if links_next:
                break
            bloco.append(tr_next)
            j += 1

        codigo, sindicatos, data_inicio, data_fim = extrair_dados_do_bloco(bloco, link_download)

        if not codigo:
            i = j
            continue

        codigos_ja_processados_na_pagina.add(codigo_tmp or codigo)
        codigos_pagina.append(codigo)

        if codigo in codigos_baixados:
            i = j
            continue

        caminho_final = baixar_e_renomear(driver, link_download, LIMPOS_DIR, codigo)

        if not caminho_final or not os.path.exists(caminho_final):
            i = j
            continue

        if not sindicatos or not data_inicio or not data_fim:
            try:
                os.remove(caminho_final)
            except OSError:
                pass
        else:
            df_registros = registrar_download(
                df_registros, codigo, sindicatos, data_inicio, data_fim
            )
            codigos_baixados.add(str(codigo))
            novos_registros_pagina += 1

        i = j

    if novos_registros_pagina:
        print(f"    {novos_registros_pagina} registro(s) novo(s) registrados nesta pagina.")

    return sorted(set(codigos_pagina)), df_registros, codigos_baixados


# ================== ESPERAR RESULTADOS ==================

def esperar_resultados(driver, timeout=TIMEOUT_RESULTADOS):
    limite = time.time() + timeout
    while time.time() < limite:
        try:
            if esta_processando(driver):
                time.sleep(1)
                continue

            links = driver.find_elements(
                By.XPATH,
                "//table[.//a[contains(@onclick,'fDownload')]]//a[contains(@onclick,'fDownload')]"
            )
            if links:
                return "tem"

            zero = driver.find_elements(
                By.XPATH,
                "//*[contains(normalize-space(.), 'Resultado: 0 Instrumento')]"
            )
            if zero:
                return "zero"
        except StaleElementReferenceException:
            pass
        except Exception:
            pass

        time.sleep(1)

    return "timeout"


# ================== PROCESSAR UM CNPJ ==================

def processar_cnpj(cnpj: str, df_registros: pd.DataFrame, codigos_baixados: set):
    driver = get_firefox_driver(LIMPOS_DIR)
    wait = WebDriverWait(driver, 60)

    try:
        driver.get(URL)
        time.sleep(2)

        chk_cnpj = wait.until(EC.element_to_be_clickable((By.ID, "chkNRCNPJ")))
        driver.execute_script("arguments[0].click();", chk_cnpj)
        time.sleep(0.5)

        driver.execute_script(f"document.getElementById('txtNRCNPJ').value = '{cnpj}';")
        try:
            driver.execute_script("$('#txtNRCNPJ').trigger('change');")
        except Exception:
            pass
        time.sleep(0.5)

        chk_vig = wait.until(EC.element_to_be_clickable((By.ID, "chkVigencia")))
        driver.execute_script("arguments[0].click();", chk_vig)
        time.sleep(0.5)

        sel_vig = wait.until(EC.presence_of_element_located((By.ID, "cboSTVigencia")))
        driver.execute_script("arguments[0].value = '2';", sel_vig)
        try:
            driver.execute_script("$('#cboSTVigencia').trigger('change');")
        except Exception:
            pass
        time.sleep(0.5)

        driver.execute_script(
            f"document.getElementById('txtDTInicioVigencia').value='{DT_INICIO_PADRAO}';"
        )
        driver.execute_script(
            f"document.getElementById('txtDTFimVigencia').value='{DT_FIM_PADRAO}';"
        )
        time.sleep(0.5)

        btn_pesquisar = wait.until(EC.element_to_be_clickable((By.ID, "btnPesquisar")))
        driver.execute_script("arguments[0].click();", btn_pesquisar)

        status_busca = esperar_resultados(driver, timeout=TIMEOUT_RESULTADOS)
        if status_busca == "zero":
            print("  Resultado: 0 Instrumento(s) para este CNPJ.")
            return df_registros, codigos_baixados
        elif status_busca == "timeout":
            print("  Timeout esperando resultados para este CNPJ.")
            return df_registros, codigos_baixados

        paginas_visitadas = set()

        while True:
            pag_atual, total_paginas = obter_paginacao_h2(driver)
            if pag_atual is None:
                pag_atual = len(paginas_visitadas) + 1

            if pag_atual in paginas_visitadas:
                break

            paginas_visitadas.add(pag_atual)

            if total_paginas is not None:
                print(f"  Pagina {pag_atual} de {total_paginas}")
            else:
                print(f"  Pagina {pag_atual}")

            codigos_pagina, df_registros, codigos_baixados = processar_pagina_atual(
                driver,
                df_registros,
                codigos_baixados
            )

            if not codigos_pagina:
                break

            if total_paginas is not None and pag_atual >= total_paginas:
                break

            if not tentar_ir_para_proxima_pagina(driver, codigos_pagina):
                break

    except Exception as e:
        print(f"  Erro ao processar CNPJ {cnpj}: {e}")
    finally:
        driver.quit()

    return df_registros, codigos_baixados


def main():
    parser = argparse.ArgumentParser(description="Worker MTE por lista de CNPJs.")
    parser.add_argument("--entrada", required=True, help="Caminho do Excel com CNPJs desta parte.")
    parser.add_argument(
        "--saida",
        required=True,
        help="Caminho do Excel de registros desta parte (sera criado/atualizado).",
    )
    parser.add_argument(
        "--id",
        type=int,
        default=1,
        help="ID do worker (apenas para log).",
    )
    args = parser.parse_args()

    entrada_cnpjs = os.path.abspath(args.entrada)
    saida_registros = os.path.abspath(args.saida)
    worker_id = args.id

    print(f"\n=== Worker {worker_id} iniciado ===")
    print(f"Worker {worker_id}: lendo CNPJs de {entrada_cnpjs}")

    df_cnpjs = pd.read_excel(entrada_cnpjs, dtype=str)
    lista_cnpjs = df_cnpjs.iloc[:, 0].dropna().astype(str).tolist()

    print(f"Worker {worker_id}: quantidade de CNPJs carregados = {len(lista_cnpjs)}")

    if not lista_cnpjs:
        print(f"Worker {worker_id}: NENHUM CNPJ para processar. Encerrando.")
        return

    if os.path.exists(saida_registros):
        df_registros = pd.read_excel(saida_registros, dtype=str)
    else:
        df_registros = pd.DataFrame(columns=["codigo", "sindicatos", "data_inicio", "data_fim"])

    for col in ["codigo", "sindicatos", "data_inicio", "data_fim"]:
        if col not in df_registros.columns:
            df_registros[col] = ""
    for col in df_registros.columns:
        df_registros[col] = df_registros[col].astype(str)

    codigos_baixados = set(df_registros["codigo"].astype(str).tolist())

    total = len(lista_cnpjs)
    for idx, cnpj in enumerate(lista_cnpjs, start=1):
        print(f"\nWorker {worker_id}: CNPJ {idx}/{total} -> {cnpj}")
        df_registros, codigos_baixados = processar_cnpj(cnpj, df_registros, codigos_baixados)
        df_registros.to_excel(saida_registros, index=False)

    print(f"\nWorker {worker_id} finalizado. Registros nesta parte: {len(df_registros)}")

if __name__ == "__main__":
    main()
