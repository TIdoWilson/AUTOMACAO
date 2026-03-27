import os
import sys
import time
import csv
import base64
import gzip
import urllib3
from datetime import datetime

from requests_pkcs12 import post
from requests.exceptions import HTTPError, RequestException
from lxml import etree

# Para janelas de seleção de arquivo/pasta
import tkinter as tk
from tkinter import filedialog, messagebox

# Desabilita o aviso de "InsecureRequestWarning" porque estamos usando verify=False
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ========================
# CONFIGURAÇÕES BÁSICAS
# ========================

# WebService NFeDistribuicaoDFe - Ambiente Nacional (produção)
WS_URL = "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"

# Ambiente: 1 = Produção, 2 = Homologação
TP_AMB = "1"

# CNPJ do INTERESSADO (destinatário dos documentos)
# <<< SUBSTITUA PELO CNPJ DA PLASTBEL >>>
CNPJ_INTERESSADO = "81472243000124"

# Certificado A1 em .pfx
CERT_PFX_PATH = r"W:\DOCUMENTOS ESCRITORIO\CERTIFICADOS DIGITAL\PLASTBEL LTDA (Plast1234).pfx"
CERT_PASSWORD = "Plast1234"

# Intervalo entre requisições (segundos) -> 2 requisições por segundo => 0.5s
INTERVALO_SEGUNDOS = 0.5

# Códigos de status da NF-e (retDistDFeInt) que indicam "documentos encontrados"
CSTAT_OK_DOCS = {"138"}   # Documentos localizados
CSTAT_NENHUM_DOC = {"137"}  # Nenhum documento localizado

# Se em algum momento você quiser tratar consumo indevido no DF-e:
IP_BLOCK_CSTATS = {"656"}  # não é comum aqui, mas deixamos preparado


# ========================
# TEMPLATES / FUNÇÕES
# ========================

# SOAP para NFeDistribuicaoDFe, consulta por chave (consChNFe)
SOAP_TEMPLATE = """<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                 xmlns:xsd="http://www.w3.org/2001/XMLSchema"
                 xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>
    <nfeDistDFeInteresse xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe">
      <nfeDadosMsg>
        <distDFeInt xmlns="http://www.portalfiscal.inf.br/nfe" versao="1.01">
          <tpAmb>{tp_amb}</tpAmb>
          <CNPJ>{cnpj}</CNPJ>
          <consChNFe>
            <chNFe>{chave}</chNFe>
          </consChNFe>
        </distDFeInt>
      </nfeDadosMsg>
    </nfeDistDFeInteresse>
  </soap12:Body>
</soap12:Envelope>
"""


def montar_soap(chave: str) -> str:
    return SOAP_TEMPLATE.format(tp_amb=TP_AMB, cnpj=CNPJ_INTERESSADO, chave=chave)


def consultar_dfe(chave: str):
    """
    Chama o NFeDistribuicaoDFe com consChNFe.
    Retorna o objeto Response.
    """
    xml_envio = montar_soap(chave)

    headers = {
        "Content-Type": "application/soap+xml; charset=utf-8",
        "SOAPAction": "http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe/nfeDistDFeInteresse",
    }

    resp = post(
        WS_URL,
        data=xml_envio.encode("utf-8"),
        headers=headers,
        pkcs12_filename=CERT_PFX_PATH,
        pkcs12_password=CERT_PASSWORD,
        timeout=30,
        verify=False,  # por causa da cadeia SSL na sua rede
    )

    return resp


def descompactar_doczip(conteudo_base64: str) -> bytes:
    """
    docZip vem BASE64 + GZIP.
    Retorna os bytes do XML interno.
    """
    comprimido = base64.b64decode(conteudo_base64)
    xml_bytes = gzip.decompress(comprimido)
    return xml_bytes


def processar_resposta(chave: str, resposta_soap: bytes, pasta_saida: str):
    """
    Processa a resposta do NFeDistribuicaoDFe.
    Retorna um dicionário de log com:
      status: "OK", "AVISO", "ERRO", "IP_BLOCK"
      cStat, xMotivo, detalhe, caminho_xml
    Pode salvar 0, 1 ou mais XMLs por chave (no geral deve vir no máximo 1).
    """
    log_entry = {
        "chave": chave,
        "status": "ERRO",
        "cStat": "",
        "xMotivo": "",
        "detalhe": "",
        "caminho_xml": "",
    }

    try:
        root = etree.fromstring(resposta_soap)

        ns = {
            "soap": "http://www.w3.org/2003/05/soap-envelope",
            "ws": "http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe",
            "nfe": "http://www.portalfiscal.inf.br/nfe",
        }

        ret = root.find(".//nfe:retDistDFeInt", ns)
        if ret is None:
            log_entry["detalhe"] = "retDistDFeInt não encontrado na resposta."
            print(f"[{chave}] ERRO: retDistDFeInt não encontrado.")
            return log_entry

        cStat = ret.findtext("nfe:cStat", namespaces=ns) or ""
        xMotivo = ret.findtext("nfe:xMotivo", namespaces=ns) or ""

        log_entry["cStat"] = cStat
        log_entry["xMotivo"] = xMotivo

        # Possível consumo indevido / bloqueio (se algum dia aparecer)
        if cStat in IP_BLOCK_CSTATS:
            log_entry["status"] = "IP_BLOCK"
            log_entry["detalhe"] = "Possível bloqueio / consumo indevido (cStat DF-e)."
            print(f"[{chave}] POSSÍVEL BLOQUEIO: cStat={cStat}, xMotivo={xMotivo}")
            return log_entry

        # Nenhum documento localizado
        if cStat in CSTAT_NENHUM_DOC:
            log_entry["status"] = "AVISO"
            log_entry["detalhe"] = "Nenhum documento localizado para esta chave (DF-e)."
            print(f"[{chave}] AVISO: Nenhum documento localizado (cStat={cStat}, xMotivo={xMotivo}).")
            return log_entry

        # Documentos encontrados
        if cStat in CSTAT_OK_DOCS:
            # Lote com docZip
            docs = ret.findall(".//nfe:docZip", ns)
            if not docs:
                log_entry["status"] = "AVISO"
                log_entry["detalhe"] = "cStat indica documentos, mas nenhum docZip retornado."
                print(f"[{chave}] AVISO: cStat 138 mas sem docZip.")
                return log_entry

            caminhos_xml = []
            for idx, doc in enumerate(docs, start=1):
                schema = doc.get("schema", "")
                nsu = doc.get("NSU", "")
                conteudo = doc.text or ""

                try:
                    xml_bytes = descompactar_doczip(conteudo)
                    # Tentamos descobrir o nome do arquivo pelo root
                    try:
                        doc_root = etree.fromstring(xml_bytes)
                        root_tag = etree.QName(doc_root.tag).localname
                    except Exception:
                        root_tag = "Desconhecido"

                    nome_arquivo = f"{chave}-NSU{nsu}-{root_tag}-{idx}.xml"
                    caminho_completo = os.path.join(pasta_saida, nome_arquivo)
                    with open(caminho_completo, "wb") as f:
                        f.write(xml_bytes)

                    caminhos_xml.append(caminho_completo)
                    print(f"[{chave}] XML DF-e salvo em {caminho_completo} (schema={schema})")
                except Exception as e:
                    print(f"[{chave}] ERRO ao descompactar/salvar docZip (NSU={nsu}): {e}")

            if caminhos_xml:
                log_entry["status"] = "OK"
                log_entry["detalhe"] = f"{len(caminhos_xml)} XML(s) salvo(s) via DF-e."
                # guardar só o primeiro no log (apenas referência)
                log_entry["caminho_xml"] = caminhos_xml[0]
            else:
                log_entry["status"] = "ERRO"
                log_entry["detalhe"] = "Não foi possível salvar nenhum XML, embora houvesse docZip."

            return log_entry

        # Qualquer outro cStat -> erro
        log_entry["status"] = "ERRO"
        log_entry["detalhe"] = "cStat não indica sucesso nem 'nenhum doc'."
        print(f"[{chave}] ERRO DF-e: cStat={cStat}, xMotivo={xMotivo}")
        return log_entry

    except Exception as e:
        log_entry["status"] = "ERRO"
        log_entry["detalhe"] = f"Falha ao processar XML DF-e: {e}"
        print(f"[{chave}] ERRO ao processar XML DF-e: {e}")
        return log_entry


def ler_chaves_de_arquivo(caminho_arquivo: str) -> list:
    if not os.path.isfile(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

    chaves = []
    with open(caminho_arquivo, "r", encoding="utf-8") as f:
        for linha in f:
            chave = linha.strip().replace("\ufeff", "")
            if not chave:
                continue
            if len(chave) != 44 or not chave.isdigit():
                raise ValueError(f"Chave inválida encontrada no arquivo: '{chave}' (deve ter 44 dígitos numéricos)")
            chaves.append(chave)

    if not chaves:
        raise ValueError("Nenhuma chave válida encontrada no arquivo.")

    return chaves


def solicitar_caminhos_ao_usuario():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    messagebox.showinfo("Seleção de arquivo", "Selecione o arquivo .txt com as chaves de acesso (uma por linha).")
    caminho_txt = filedialog.askopenfilename(
        title="Selecione o arquivo .txt com as chaves de acesso",
        filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")],
    )

    if not caminho_txt:
        root.destroy()
        raise ValueError("Nenhum arquivo .txt selecionado.")

    messagebox.showinfo("Seleção de pasta", "Selecione a pasta onde deseja salvar os XML das NF-e.")
    pasta_saida = filedialog.askdirectory(
        title="Selecione a pasta de saída para os XML",
    )

    root.destroy()

    if not pasta_saida:
        raise ValueError("Nenhuma pasta de saída selecionada.")

    return caminho_txt, pasta_saida


def criar_arquivo_log(pasta_saida: str) -> str:
    agora = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_log = f"log_distribuicao_dfe_{agora}.csv"
    caminho_log = os.path.join(pasta_saida, nome_log)

    with open(caminho_log, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["data_hora", "chave", "status", "cStat", "xMotivo", "detalhe", "caminho_xml"])

    return caminho_log


def escrever_log(caminho_log: str, entry: dict):
    with open(caminho_log, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            entry.get("chave", ""),
            entry.get("status", ""),
            entry.get("cStat", ""),
            entry.get("xMotivo", ""),
            entry.get("detalhe", ""),
            entry.get("caminho_xml", ""),
        ])


def main():
    # Checagem básica do CNPJ
    if not (len(CNPJ_INTERESSADO) == 14 and CNPJ_INTERESSADO.isdigit()):
        print("ERRO: Preencha corretamente o CNPJ_INTERESSADO (14 dígitos numéricos).")
        sys.exit(1)

    try:
        caminho_txt, pasta_saida = solicitar_caminhos_ao_usuario()
        chaves = ler_chaves_de_arquivo(caminho_txt)
    except Exception as e:
        print(f"\nErro na configuração inicial: {e}")
        sys.exit(1)

    print(f"\nArquivo de chaves: {caminho_txt}")
    print(f"Pasta de saída: {pasta_saida}")
    print(f"Total de chaves carregadas: {len(chaves)}")
    print(f"CNPJ interessado: {CNPJ_INTERESSADO}")
    print(f"Limite configurado: 2 requisições por segundo (intervalo de {INTERVALO_SEGUNDOS} s)")

    caminho_log = criar_arquivo_log(pasta_saida)
    print(f"Log das consultas: {caminho_log}")

    input("\nPressione ENTER para iniciar as consultas DF-e...")

    for i, chave in enumerate(chaves, start=1):
        print(f"\n({i}/{len(chaves)}) Consultando DF-e para chave {chave} ...")

        log_entry = {
            "chave": chave,
            "status": "ERRO",
            "cStat": "",
            "xMotivo": "",
            "detalhe": "",
            "caminho_xml": "",
        }

        try:
            resp = consultar_dfe(chave)
            try:
                resp.raise_for_status()
            except HTTPError as http_err:
                status_code = http_err.response.status_code
                if status_code in (403, 429, 503):
                    log_entry["status"] = "IP_BLOCK"
                    log_entry["detalhe"] = f"HTTP {status_code} - possível bloqueio de IP."
                    print(f"[{chave}] POSSÍVEL BLOQUEIO DE IP: HTTP {status_code}. Encerrando.")
                    escrever_log(caminho_log, log_entry)
                    break
                else:
                    log_entry["status"] = "ERRO"
                    log_entry["detalhe"] = f"HTTP Error {status_code}: {http_err}"
                    print(f"[{chave}] HTTP ERROR {status_code}: {http_err}")
                    escrever_log(caminho_log, log_entry)
                    time.sleep(INTERVALO_SEGUNDOS)
                    continue

            resposta_bytes = resp.content
            log_entry = processar_resposta(chave, resposta_bytes, pasta_saida)
            escrever_log(caminho_log, log_entry)

            if log_entry["status"] == "IP_BLOCK":
                print(f"[{chave}] POSSÍVEL BLOQUEIO DE IP (cStat={log_entry['cStat']}). Encerrando.")
                break

        except RequestException as req_err:
            log_entry["status"] = "ERRO"
            log_entry["detalhe"] = f"Erro de rede/requests: {req_err}"
            print(f"[{chave}] ERRO de rede/requests: {req_err}")
            escrever_log(caminho_log, log_entry)

        except Exception as e:
            log_entry["status"] = "ERRO"
            log_entry["detalhe"] = f"Erro inesperado: {e}"
            print(f"[{chave}] ERRO inesperado: {e}")
            escrever_log(caminho_log, log_entry)

        time.sleep(INTERVALO_SEGUNDOS)

    print("\nProcessamento DF-e concluído (fim da lista ou possível bloqueio de IP detectado).")
    print(f"Log completo em: {caminho_log}")


if __name__ == "__main__":
    main()
