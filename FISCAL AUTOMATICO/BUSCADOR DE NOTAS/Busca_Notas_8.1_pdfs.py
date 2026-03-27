import os
import time
from datetime import datetime
import subprocess
import configparser
import unicodedata
import re
import sys
import shutil
import requests
import urllib3
from playwright.sync_api import sync_playwright
from concurrent.futures import ThreadPoolExecutor, as_completed

# Desativa warning de certificado inseguro (apenas para este script)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================================
# CONFIGURAÇÕES
# ============================================================

# Pasta base dos downloads
base_download_dir = r"W:\XML PREFEITURA"

# Caminho opcional do script principal a ser chamado no final
CAMINHO_SEM_MVTO = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\BUSCADOR DE NOTAS\emitir_declaracao.py"
PERFIL_CHROME = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\BUSCADOR DE NOTAS\perfil_chrome_esnfs"

ESNFS_BASE_URL = "https://www.esnfs.com.br"
URL_CONSULTA_NOTAS = f"{ESNFS_BASE_URL}/nfsconsultanota.list.logic"
URL_DOWNLOAD_NOTAS = f"{ESNFS_BASE_URL}/nfsconsultanota.download.logic"

URL_RELATORIO_LIST = f"{ESNFS_BASE_URL}/nfsrelapuracaoiss.list.logic"
URL_RELATORIO_DOWNLOAD_PDF = f"{ESNFS_BASE_URL}/nfsrelapuracaoiss.downloadPdf.logic"

HEADERS_HTTP = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Content-Type": "application/x-www-form-urlencoded",
    "Origin": ESNFS_BASE_URL,
    "Referer": URL_CONSULTA_NOTAS,
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36",
}

HEADERS_RELATORIO = {
    "Accept": HEADERS_HTTP["Accept"],
    "Content-Type": "application/x-www-form-urlencoded",
    "Origin": ESNFS_BASE_URL,
    "Referer": URL_RELATORIO_LIST,
    "User-Agent": HEADERS_HTTP["User-Agent"],
}

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

# Quantidade de threads para as requisições HTTP
MAX_WORKERS = 2

# ============================================================
# LOG SIMPLES
# ============================================================

LOG_ARQUIVO = None

def _resolver_caminho_log():
    global LOG_ARQUIVO
    if LOG_ARQUIVO:
        return LOG_ARQUIVO
    try:
        base_dir = base_download_dir if base_download_dir else os.getcwd()
        LOG_ARQUIVO = os.path.join(base_dir, "log_operacoes.txt")
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

# ============================================================
# UTILS
# ============================================================
def salvar_captura_de_tela(pagina, caminho, nome_arquivo):
    os.makedirs(caminho, exist_ok=True)
    caminho_arquivo = os.path.join(caminho, nome_arquivo)
    try:
        pagina.screenshot(path=caminho_arquivo, full_page=True)
        print(f"📸 Captura de tela salva em: {caminho_arquivo}")
    except Exception as e:
        print(f"❗ Erro ao salvar captura de tela: {e}")

# ============================================================
# HTTP (após login)
# ============================================================
def obter_perfil_firefox_padrao() -> str:
    """
    Localiza e retorna o caminho do *perfil padrão do Firefox* no Windows.

    Prioridade:
      1) installs.ini (Firefox mais recente)
      2) profiles.ini (Profile com Default=1)
      3) primeiro Profile disponível (fallback)

    Observação: usar o perfil real do Firefox pode falhar se o Firefox estiver aberto,
    por conta de lock do perfil.
    """
    appdata = os.environ.get("APPDATA") or ""
    base = os.path.join(appdata, "Mozilla", "Firefox")

    installs_ini = os.path.join(base, "installs.ini")
    profiles_ini = os.path.join(base, "profiles.ini")

    if not os.path.isfile(profiles_ini) and not os.path.isfile(installs_ini):
        raise RuntimeError(f"Não encontrei installs.ini/profiles.ini em: {base}")

    # 1) installs.ini (mais confiável para 'default profile')
    if os.path.isfile(installs_ini):
        cp = configparser.RawConfigParser()
        cp.read(installs_ini, encoding="utf-8")
        for sec in cp.sections():
            if sec.lower().startswith("install") and cp.has_option(sec, "Default"):
                rel = (cp.get(sec, "Default") or "").strip()
                if rel:
                    cand = os.path.join(base, rel)
                    if os.path.isdir(cand):
                        return cand

    # 2) profiles.ini com Default=1
    if os.path.isfile(profiles_ini):
        cp2 = configparser.RawConfigParser()
        cp2.read(profiles_ini, encoding="utf-8")

        def _abs_path(sec: str) -> str:
            pth = (cp2.get(sec, "Path", fallback="") or "").strip()
            if not pth:
                return ""
            is_rel = (cp2.get(sec, "IsRelative", fallback="1") or "1").strip() == "1"
            return os.path.join(base, pth) if is_rel else pth

        for sec in cp2.sections():
            if sec.lower().startswith("profile"):
                is_default = (cp2.get(sec, "Default", fallback="0") or "0").strip() == "1"
                if is_default:
                    cand = _abs_path(sec)
                    if cand and os.path.isdir(cand):
                        return cand

        # 3) fallback: primeiro perfil válido
        for sec in cp2.sections():
            if sec.lower().startswith("profile"):
                cand = _abs_path(sec)
                if cand and os.path.isdir(cand):
                    return cand

    raise RuntimeError(f"Não foi possível localizar um perfil válido do Firefox em: {base}")


def obter_perfil_firefox() -> str:
    """
    Prioriza o perfil dedicado do script, que ja costuma ter as excecoes
    e preferencias de certificado salvas.
    """
    if os.path.isdir(PERFIL_FIREFOX_CERT):
        return PERFIL_FIREFOX_CERT
    return obter_perfil_firefox_padrao()


def criar_sessao_http_com_cookies(cookies_dict: dict) -> requests.Session:
    """
    Cria uma requests.Session com verify=False e os cookies do login.
    Usada por cada thread.
    """
    s = requests.Session()
    s.verify = False
    for name, value in cookies_dict.items():
        s.cookies.set(name, value, domain="www.esnfs.com.br")
    return s


def corrigir_mojibake(texto: str) -> str:
    """Corrige casos comuns de UTF-8 decodificado como Latin-1/Windows-1252 (ex.: 'INDÃºSTRIA')."""
    if not isinstance(texto, str):
        texto = str(texto or "")
    if "Ã" in texto or "Â" in texto:
        try:
            return texto.encode("latin1").decode("utf-8")
        except UnicodeError:
            return texto
    return texto

def limpar_nome(nome: str) -> str:
    if not isinstance(nome, str):
        nome = str(nome or "")

    # Corrige mojibake antes de normalizar (preserva acentos corretamente)
    nome = corrigir_mojibake(nome)

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


def carregar_prestadores(session_http: requests.Session):
    """
    Faz GET na tela de consulta e extrai os <option> do
    select[name="parametrosTela.idPessoa"] via regex, porque o HTML é malformado.
    Retorna lista de (value, texto_original, texto_limpo).
    """
    resp = session_http.get(URL_CONSULTA_NOTAS, headers=HEADERS_HTTP, timeout=30)
    resp.raise_for_status()
    html = resp.text

    # Corrige casos comuns de UTF-8 decodificado como Latin-1/Windows-1252
    html = corrigir_mojibake(html)

    m_sel = re.search(
        r'<select[^>]+name="parametrosTela\.idPessoa"[^>]*>(.*?)</select>',
        html,
        flags=re.I | re.S,
    )
    if not m_sel:
        raise RuntimeError("Não encontrei o <select name='parametrosTela.idPessoa'> na tela de consulta.")

    inner = m_sel.group(1)

    pattern = r'<option\s+value="([^"]*)">(.*?)(?=(<option\s+value=|</select>))'
    prestadores_info = []

    for m in re.finditer(pattern, inner, flags=re.I | re.S):
        val = (m.group(1) or "").strip()
        txt_original = (m.group(2) or "").strip()
        txt_original = re.sub(r"\s+", " ", txt_original)
        txt_original = corrigir_mojibake(txt_original)

        if not val:
            continue

        txt_limpo = limpar_nome(txt_original)
        prestadores_info.append((val, txt_original, txt_limpo))

    if not prestadores_info:
        raise RuntimeError("Nenhum <option value='...'> válido encontrado no select de prestadores.")

    print("Exemplo de prestadores parseados:")
    for v, t_orig, t_clean in prestadores_info[:5]:
        print(f"  value={v} | original={t_orig} | limpo={t_clean}")

    return prestadores_info

def montar_payload_consulta(valor_prestador: str, mes_emissao: int, ano_emissao: int, origem_texto: str) -> dict:
    origem_lower = (origem_texto or "").lower()

    if origem_lower == "emitida":
        origem_codigo = "1"
        id_pessoa = str(valor_prestador)
        id_prestador = str(valor_prestador)
        id_tomador = ""
    else:
        origem_codigo = "2"
        id_pessoa = str(valor_prestador)
        id_prestador = ""
        id_tomador = ""

    return {
        "idNotaParaReenvioDeEmail": "",
        "parametrosTela.ehAdmin": "false",
        "parametrosTela.exportarTodasAsNotas": "false",
        "emailAlternativoParaReenvioDeNota": "",
        "parametrosTela.nmPessoa": "",
        "parametrosTela.origemEmissaoNfse": origem_codigo,
        "parametrosTela.idPessoa": id_pessoa,
        "parametrosTela.idPessoaPrestador": id_prestador,
        "parametrosTela.nrDocumentoPrestador": "",
        "parametrosTela.nmPessoaPrestador": "",
        "parametrosTela.idPessoaTomador": id_tomador,
        "parametrosTela.nrDocumentoTomador": "",
        "parametrosTela.nmPessoaTomador": "",
        "parametrosTela.nrMesCompetencia": "",
        "parametrosTela.nrAnoCompetencia": "",
        "parametrosTela.nrMesCompetenciaEmissao": str(mes_emissao),
        "parametrosTela.nrAnoCompetenciaEmissao": str(ano_emissao),
        "parametrosTela.dtEmissaoInicial": "",
        "parametrosTela.dtEmissaoFinal": "",
        "parametrosTela.tpConsultaNfs": "0",
        "parametrosTela.nrNfs": "",
        "parametrosTela.nrRps": "",
        "parametrosTela.nrSerieRps": "",
    }

def montar_payload_download(valor_prestador: str, mes_emissao: int, ano_emissao: int, origem_texto: str) -> dict:
    origem_lower = (origem_texto or "").lower()

    if origem_lower == "emitida":
        origem_codigo = "1"
        id_pessoa = str(valor_prestador)
        id_prestador = str(valor_prestador)
        id_tomador = ""
    else:
        origem_codigo = "2"
        id_pessoa = str(valor_prestador)
        id_prestador = ""
        id_tomador = str(valor_prestador)

    return {
        "idNotaParaReenvioDeEmail": "",
        "parametrosTela.ehAdmin": "false",
        "parametrosTela.exportarTodasAsNotas": "false",
        "emailAlternativoParaReenvioDeNota": "",
        "parametrosTela.nmPessoa": "",
        "parametrosTela.origemEmissaoNfse": origem_codigo,
        "parametrosTela.idPessoa": id_pessoa,
        "parametrosTela.idPessoaPrestador": id_prestador,
        "parametrosTela.nrDocumentoPrestador": "",
        "parametrosTela.nmPessoaPrestador": "",
        "parametrosTela.idPessoaTomador": id_tomador,
        "parametrosTela.nrDocumentoTomador": "",
        "parametrosTela.nmPessoaTomador": "",
        "parametrosTela.nrMesCompetencia": "",
        "parametrosTela.nrAnoCompetencia": "",
        "parametrosTela.nrMesCompetenciaEmissao": str(mes_emissao),
        "parametrosTela.nrAnoCompetenciaEmissao": str(ano_emissao),
        "parametrosTela.dtEmissaoInicial": "",
        "parametrosTela.dtEmissaoFinal": "",
        "parametrosTela.tpConsultaNfs": "0",
        "parametrosTela.nrNfs": "",
        "parametrosTela.nrRps": "",
        "parametrosTela.nrSerieRps": "",
    }


def montar_payload_relatorio_pdf(valor_prestador: str, mes_competencia: int, ano_competencia: int, tipo_origem: str) -> dict:
    """
    Monta payload do relatório de Apuração do ISS (PDF) conforme capturas de rede:
      - tpOrigemNfs: EMITIDA | RECEBIDA (uppercase)
      - demais campos conforme formulário.
    """
    origem = (tipo_origem or "").strip().upper()
    if origem not in ("EMITIDA", "RECEBIDA"):
        # aceita também "Emitida"/"Recebida"
        origem = "EMITIDA" if (tipo_origem or "").lower() == "emitida" else "RECEBIDA"

    return {
        "formulario.ehAdmin": "false",
        "formulario.idPessoa": str(valor_prestador),
        "formulario.tpOrigemNfs": origem,
        "formulario.nrInscricaoMunicipal": "",
        "formulario.nrMesCompetencia": str(mes_competencia),
        "formulario.nrAnoCompetencia": str(ano_competencia),
        "formulario.dtEmissaoInicial": "",
        "formulario.dtEmissaoFinal": "",
        "formulario.tpConsultaNfs": "0",
    }

def consultar_e_baixar_xml_http(
    cookies_dict: dict,
    nome_prestador: str,
    valor_prestador: str,
    mes_anterior: int,
    ano_ref: int,
    mes_extenso: str,
    tipo_origem: str,  # "Emitida" ou "Recebida"
):
    """
    Faz CONSULTA + DOWNLOAD via HTTP (requests) para um prestador/origem.
    NÃO faz screenshot aqui (isso ficará para a etapa sequencial depois).
    Retorna (tipo_origem, nome_prestador, valor_prestador, tem_registro, pasta_mes_ano)
    """
    sufixo = "emitido" if tipo_origem.lower() == "emitida" else "recebido"

    session_http = criar_sessao_http_com_cookies(cookies_dict)

    nome_limpo = nome_prestador.split(" - ", 1)[1] if " - " in nome_prestador else nome_prestador
    nome_limpo = limpar_nome(nome_limpo)
    pasta_cliente = os.path.join(base_download_dir, nome_limpo)
    pasta_mes_ano = os.path.join(pasta_cliente, f"{mes_anterior:02d}.{ano_ref}")
    os.makedirs(pasta_mes_ano, exist_ok=True)

    payload = montar_payload_consulta(valor_prestador, mes_anterior, ano_ref, tipo_origem)
    tem_registro = False

    try:
        resp_list = session_http.post(URL_CONSULTA_NOTAS, headers=HEADERS_HTTP, data=payload, timeout=30)
        resp_list.raise_for_status()
        html = resp_list.text

        if "Não há registros" in html or "Nao ha registros" in html:
            tem_registro = False
            print(f"ℹ️ {tipo_origem}: Sem registros para {mes_anterior:02d}/{ano_ref} - {nome_limpo}")
        elif "tabelaDinamica" in html:
            tem_registro = True
            print(f"✅ {tipo_origem}: Há registros para {mes_anterior:02d}/{ano_ref} - {nome_limpo}")
        else:
            tem_registro = False
            print(f"⚠️ {tipo_origem}: Resposta não reconhecida para {nome_limpo}. Assumindo SEM registros.")
    except Exception as e:
        print(f"⚠️ Erro ao consultar notas via HTTP ({tipo_origem}) [{nome_limpo}]: {e}")
        tem_registro = False

    if tem_registro:
        try:
            payload_down = montar_payload_download(
                valor_prestador=valor_prestador,
                mes_emissao=mes_anterior,
                ano_emissao=ano_ref,
                origem_texto=tipo_origem,
            )

            resp_down = session_http.post(
                URL_DOWNLOAD_NOTAS,
                headers=HEADERS_HTTP,
                data=payload_down,
                timeout=60,
            )
            resp_down.raise_for_status()

            xml_text = resp_down.text.strip()

            if "<listaNfs" in xml_text and "<nfs" not in xml_text:
                print(f"ℹ️ {tipo_origem}: download retornou listaNfs vazia para {mes_anterior:02d}/{ano_ref} - {nome_limpo}")
                registrar_log(
                    nome_prestador,
                    f"XML VAZIO (listaNfs) {tipo_origem.upper()} {mes_anterior:02d}/{ano_ref}",
                )
            else:
                nome_arquivo_xml = f"notas_{mes_extenso.lower()}_{ano_ref}_{sufixo}.xml"
                caminho_final_xml = os.path.join(pasta_mes_ano, nome_arquivo_xml)
                with open(caminho_final_xml, "wb") as f:
                    f.write(resp_down.content)

                print(f"✅ XML {tipo_origem} salvo em:\n{caminho_final_xml}")
                registrar_log(
                    nome_prestador,
                    f"SUCESSO XML {sufixo.upper()} {mes_anterior:02d}/{ano_ref} | {caminho_final_xml}",
                )


                # Se houve movimento para esta origem, baixa também o relatório PDF (Apuração do ISS)
                try:
                    baixar_relatorio_pdf_http(
                        cookies_dict=cookies_dict,
                        nome_prestador=nome_prestador,
                        valor_prestador=valor_prestador,
                        mes_competencia=mes_anterior,
                        ano_competencia=ano_ref,
                        mes_extenso=mes_extenso,
                        tipo_origem=tipo_origem,
                        pasta_mes_ano=pasta_mes_ano,
                    )
                except Exception as e:
                    print(f"⚠️ Erro ao baixar PDF ({tipo_origem}) [{nome_limpo}]: {e}")
        except Exception as e:
            print(f"⚠️ Falha ao baixar XML via HTTP ({tipo_origem}) [{nome_limpo}]: {e}")
            registrar_log(
                nome_prestador,
                f"ERRO XML {sufixo.upper()} {mes_anterior:02d}/{ano_ref}: {e}",
            )


            # Mesmo se o XML falhar, ainda tentamos baixar o PDF (pode estar disponível)
            try:
                baixar_relatorio_pdf_http(
                    cookies_dict=cookies_dict,
                    nome_prestador=nome_prestador,
                    valor_prestador=valor_prestador,
                    mes_competencia=mes_anterior,
                    ano_competencia=ano_ref,
                    mes_extenso=mes_extenso,
                    tipo_origem=tipo_origem,
                    pasta_mes_ano=pasta_mes_ano,
                )
            except Exception as e2:
                print(f"⚠️ Erro ao baixar PDF após falha de XML ({tipo_origem}) [{nome_limpo}]: {e2}")
    else:
        registrar_log(
            nome_prestador,
            f"SEM REGISTROS {sufixo.upper()} {mes_anterior:02d}/{ano_ref}",
        )

    return (tipo_origem, nome_prestador, valor_prestador, tem_registro, pasta_mes_ano)

# ============================================================
# PLAYWRIGHT – SCREENSHOT SEM MOVIMENTO
# ============================================================


def baixar_relatorio_pdf_http(
    cookies_dict: dict,
    nome_prestador: str,
    valor_prestador: str,
    mes_competencia: int,
    ano_competencia: int,
    mes_extenso: str,
    tipo_origem: str,  # "Emitida" ou "Recebida"
    pasta_mes_ano: str,
) -> bool:
    """
    Faz download do PDF (Apuração do ISS) via HTTP.
    Estratégia: se já houve movimento (mesma base usada para XML), chama diretamente o endpoint downloadPdf.logic
    sem precisar chamar list.logic antes.

    Retorna True se salvou um PDF válido.
    """
    sufixo = "emitido" if (tipo_origem or "").lower() == "emitida" else "recebido"
    session_http = criar_sessao_http_com_cookies(cookies_dict)

    payload = montar_payload_relatorio_pdf(valor_prestador, mes_competencia, ano_competencia, tipo_origem)
    try:
        resp = session_http.post(
            URL_RELATORIO_DOWNLOAD_PDF,
            headers=HEADERS_RELATORIO,
            data=payload,
            timeout=90,
        )
        resp.raise_for_status()

        # Alguns casos retornam HTML (erro/sessão expirada) em vez de PDF.
        content_type = (resp.headers.get("Content-Type") or "").lower()
        is_pdf = resp.content[:4] == b"%PDF" or "application/pdf" in content_type

        if not is_pdf:
            # registra amostra curta do retorno para diagnóstico, sem salvar conteúdo sensível completo
            amostra = (resp.text or "")[:200].replace("\n", " ").replace("\r", " ")
            registrar_log(
                nome_prestador,
                f"ERRO PDF {sufixo.upper()} {mes_competencia:02d}/{ano_competencia} | Retorno não-PDF (Content-Type={content_type}) | amostra={amostra}",
            )
            return False

        nome_arquivo_pdf = f"notas_{mes_extenso.lower()}_{ano_competencia}_{sufixo}.pdf"
        caminho_final_pdf = os.path.join(pasta_mes_ano, nome_arquivo_pdf)
        with open(caminho_final_pdf, "wb") as f:
            f.write(resp.content)

        print(f"✅ PDF {tipo_origem} salvo em:\n{caminho_final_pdf}")
        registrar_log(
            nome_prestador,
            f"SUCESSO PDF {sufixo.upper()} {mes_competencia:02d}/{ano_competencia} | {caminho_final_pdf}",
        )
        return True

    except Exception as e:
        print(f"⚠️ Falha ao baixar PDF via HTTP ({tipo_origem}) [{nome_prestador}]: {e}")
        registrar_log(
            nome_prestador,
            f"ERRO PDF {sufixo.upper()} {mes_competencia:02d}/{ano_competencia}: {e}",
        )
        return False

def tirar_screenshot_sem_movimento(
    pagina,
    nome_prestador: str,
    valor_prestador: str,
    mes_anterior: int,
    ano_ref: int,
    tipo_origem: str,
    pasta_mes_ano: str,
):
    """
    Usa o Playwright para ir até a tela de consulta, selecionar o prestador,
    origem, mês/ano, clicar em Pesquisar e tirar um print da tela de "sem movimento".
    Rodamos isso SEQUENCIALMENTE depois que as requisições HTTP terminarem.
    """
    sufixo = "emitido" if tipo_origem.lower() == "emitida" else "recebido"

    try:
        # Garante que estamos na tela de consulta
        try:
            pagina.click("text=NFS-E")
            pagina.click("text=Consulta")
            pagina.wait_for_load_state("domcontentloaded")
        except Exception:
            # Se já estiver na tela, só segue
            pass

        # Seleciona prestador e filtros
        pagina.wait_for_selector('select[name="parametrosTela.idPessoa"]', timeout=30000)
        prestador_select = pagina.locator('select[name="parametrosTela.idPessoa"]')
        prestador_select.select_option(value=valor_prestador)

        pagina.wait_for_selector('select[name="parametrosTela.origemEmissaoNfse"]', timeout=30000)
        pagina.select_option('select[name="parametrosTela.origemEmissaoNfse"]', label=tipo_origem)

        # Tenta usar os selects de competência (se existirem)
        try:
            pagina.wait_for_selector('select[name="parametrosTela.nrMesCompetencia"]', timeout=5000)
            pagina.select_option('select[name="parametrosTela.nrMesCompetencia"]', str(mes_anterior))
            pagina.select_option('select[name="parametrosTela.nrAnoCompetencia"]', str(ano_ref))
        except Exception:
            # se não tiver, tenta por emissão
            try:
                pagina.wait_for_selector('select[name="parametrosTela.nrMesCompetenciaEmissao"]', timeout=5000)
                pagina.select_option('select[name="parametrosTela.nrMesCompetenciaEmissao"]', str(mes_anterior))
                pagina.select_option('select[name="parametrosTela.nrAnoCompetenciaEmissao"]', str(ano_ref))
            except Exception:
                pass

        # Clica em Pesquisar
        pagina.click("text=Pesquisar")
        pagina.wait_for_load_state("domcontentloaded")
        time.sleep(1)

        # Agora tira o print
        nome_arquivo = f"sem_movimento_{mes_anterior:02d}.{ano_ref}_{sufixo}.png"
        salvar_captura_de_tela(pagina, pasta_mes_ano, nome_arquivo)

    except Exception as e:
        print(f"⚠️ Erro ao tirar screenshot sem movimento ({tipo_origem}) [{nome_prestador}]: {e}")

# ================== FLUXO PRINCIPAL ==================

with sync_playwright() as p:
    # Usa diretamente o perfil padrão do Firefox (caminho padrão do Windows)
    
    os.makedirs(PERFIL_CHROME, exist_ok=True)
    contexto = p.chromium.launch_persistent_context(
        user_data_dir=PERFIL_CHROME,
        channel="chrome",
        headless=False,
        accept_downloads=True,
        ignore_https_errors=True,
        args=["--ignore-certificate-errors"],
    )
    print(f"Perfil Chrome em uso: {PERFIL_CHROME}")
    pagina = contexto.new_page()

    # Acesso ao site
    pagina.goto("https://www.esnfs.com.br/?e=35")
    time.sleep(3)

    # Fecha modal "Fechar" se aparecer
    try:
        fechar_btn = pagina.locator("div.modal-footer button[data-dismiss='modal'], button:has-text('Fechar')").first
        fechar_btn.wait_for(state="visible", timeout=1000)
        fechar_btn.click()
    except Exception:
        pass


    # Login por CERTIFICADO DIGITAL
    botao_cert = pagina.locator('button[onclick*="useDigitalCertificate=true"]')
    botao_cert.wait_for(state="visible", timeout=1000)
    botao_cert.click()


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

    pagina.goto("https://www.esnfs.com.br/nfsconsultanota.load.logic")

    # ========== MONTA COOKIES PARA HTTP ==========
    cookies_list = contexto.cookies(ESNFS_BASE_URL)
    cookies_dict = {
        c["name"]: c["value"]
        for c in cookies_list
        if "esnfs.com.br" in c.get("domain", "")
    }

    # usa uma sessão só para carregar prestadores
    session_http_main = criar_sessao_http_com_cookies(cookies_dict)
    prestadores_info = carregar_prestadores(session_http_main)
    if not prestadores_info:
        raise RuntimeError("Nenhum prestador encontrado na tela de consulta.")

    print(f"📄 Prestadores carregados: {len(prestadores_info)}")

    # ========== CONSULTAS/DOWNLOADS EM PARALELO ==========
    resultados = []  # (tipo_origem, nome_prestador, valor_prestador, tem_registro, pasta_mes_ano)
    futures = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        for valor_prestador, nome_original, nome_limpo in prestadores_info:
            nome_prestador = nome_limpo
            # Emitida
            futures.append(
                executor.submit(
                    consultar_e_baixar_xml_http,
                    cookies_dict,
                    nome_prestador,
                    valor_prestador,
                    mes_anterior,
                    ano_ref,
                    mes_extenso,
                    "Emitida",
                )
            )
            # Recebida
            futures.append(
                executor.submit(
                    consultar_e_baixar_xml_http,
                    cookies_dict,
                    nome_prestador,
                    valor_prestador,
                    mes_anterior,
                    ano_ref,
                    mes_extenso,
                    "Recebida",
                )
            )

        total_futures = len(futures)
        for i, f in enumerate(as_completed(futures), start=1):
            try:
                res = f.result()
                resultados.append(res)
            except Exception as e:
                print(f"⚠️ Erro em uma tarefa HTTP: {e}")
            if i % 20 == 0 or i == total_futures:
                print(f"… {i}/{total_futures} operações HTTP concluídas")

    # ========== SCREENSHOT SEM MOVIMENTO (SEQUENCIAL) ==========
    for tipo_origem, nome_prestador, valor_prestador, tem_registro, pasta_mes_ano in resultados:
        if not tem_registro:
            tirar_screenshot_sem_movimento(
                pagina,
                nome_prestador,
                valor_prestador,
                mes_anterior,
                ano_ref,
                tipo_origem,
                pasta_mes_ano,
            )

    contexto.close()

    # Se quiser manter a chamada do importador:
    if os.path.isfile(CAMINHO_SEM_MVTO):
        print("\n🚀 Executando principal.py...")
        subprocess.run(["python", CAMINHO_SEM_MVTO], check=True)
    else:
        print(f"⚠️ principal.py não encontrado em: {CAMINHO_SEM_MVTO}")
