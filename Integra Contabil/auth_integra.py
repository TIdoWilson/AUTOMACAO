# auth_integra.py
import os
import base64
import logging
import site
import json
import re
from pathlib import Path
from typing import Optional, Dict, Any, Tuple

# Adicionar site-packages do usuário para compatibilidade com venvs
site.addsitedir(site.getusersitepackages())

from dotenv import load_dotenv
from requests_pkcs12 import post as pkcs12_post
import requests
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# Constantes para Integra Contador
BASE_URL = "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1"
ID_SERVICOS_EMISSAO: Dict[str, str] = {
    "PARCSN": "GERARDAS161",
    "PARCSN-ESP": "GERARDAS171",
    "PERTSN": "GERARDAS181",
    "RELPSN": "GERARDAS191",
    "PGFN": "PGFN",  # Assuming PGFN uses "PGFN" as idServico
}

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(BASE_DIR / "auth_integra.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Logger separado para cobranças
charges_logger = logging.getLogger('charges')
charges_logger.setLevel(logging.INFO)
charges_handler = logging.FileHandler(BASE_DIR / "charges.log", encoding='utf-8')
charges_formatter = logging.Formatter('%(asctime)s - %(message)s')
charges_handler.setFormatter(charges_formatter)
charges_logger.addHandler(charges_handler)


def log_cobranca(descricao: str, valor: Optional[float] = None, detalhes: Optional[str] = None):
    """
    Registra uma possível cobrança no log de cobranças.

    Args:
        descricao: Descrição da cobrança (ex.: "Autenticação Serpro", "Consulta CNPJ").
        valor: Valor estimado da cobrança (opcional).
        detalhes: Detalhes adicionais (opcional).
    """
    mensagem = f"COBRANÇA POSSÍVEL: {descricao}"
    if valor is not None:
        mensagem += f" - Valor estimado: R$ {valor:.2f}"
    if detalhes:
        mensagem += f" - Detalhes: {detalhes}"
    charges_logger.info(mensagem)


def _limpar_caminho_pfx(raw: Optional[str]) -> str:
    """
    Limpa problemas comuns no SERPRO_PFX_PATH:
    - remove espaços extras
    - remove '=' no início (caso venha de fórmula de Excel)
    - remove aspas simples/dobras em volta
    """
    if raw is None:
        return ""

    s = raw.strip()

    # remove '=' do começo (ex: '=W:...' ou '==W:...')
    while s.startswith("="):
        s = s[1:].lstrip()

    # remove aspas ao redor
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        s = s[1:-1].strip()

    return s

def obter_tokens() -> tuple[str, str]:
    """
    Autentica no Serpro (Integra Contador) e devolve (access_token, jwt_token).

    Requer no .env (na pasta mãe):
        SERPRO_CONSUMER_KEY
        SERPRO_CONSUMER_SECRET
        SERPRO_PFX_PATH
        SERPRO_PFX_PASSWORD
    """
    logger.info("Iniciando autenticação no Serpro.")

    consumer_key = os.getenv("SERPRO_CONSUMER_KEY")
    consumer_secret = os.getenv("SERPRO_CONSUMER_SECRET")
    pfx_path_raw = os.getenv("SERPRO_PFX_PATH")
    pfx_password = os.getenv("SERPRO_PFX_PASSWORD")

    if not all([consumer_key, consumer_secret, pfx_path_raw, pfx_password]):
        logger.error("Variáveis de ambiente faltando no .env.")
        raise RuntimeError(
            "Faltam variáveis no .env: "
            "SERPRO_CONSUMER_KEY, SERPRO_CONSUMER_SECRET, "
            "SERPRO_PFX_PATH, SERPRO_PFX_PASSWORD."
        )

    pfx_path = _limpar_caminho_pfx(pfx_path_raw)
    pfx_file = Path(pfx_path)

    if not pfx_file.exists():
        logger.error(f"Arquivo PFX não encontrado: {str(pfx_file)}")
        raise FileNotFoundError(
            "Caminho do certificado PFX não encontrado.\n"
            f"  SERPRO_PFX_PATH (bruto): {pfx_path_raw!r}\n"
            f"  Caminho após limpeza:   {str(pfx_file)!r}"
        )

    basic = base64.b64encode(f"{consumer_key}:{consumer_secret}".encode()).decode()

    url = "https://autenticacao.sapi.serpro.gov.br/authenticate"

    headers = {
        "Authorization": f"Basic {basic}",
        "role-type": "TERCEIROS",
        "content-type": "application/x-www-form-urlencoded",
    }

    data = {"grant_type": "client_credentials"}

    try:
        logger.info("Enviando requisição de autenticação para Serpro.")
        resp = pkcs12_post(
            url,
            data=data,
            headers=headers,
            pkcs12_filename=str(pfx_file),
            pkcs12_password=pfx_password,
            timeout=30  # Timeout de 30 segundos
        )
        resp.raise_for_status()
        j = resp.json()

        access_token = j["access_token"]
        jwt_token = j["jwt_token"]

        logger.info("Autenticação bem-sucedida. Tokens obtidos.")
        # Registrar possível cobrança pela autenticação
        log_cobranca("Autenticação Serpro (Integra Contador)", valor=0.00, detalhes="Obtenção de tokens OAuth2")

        return access_token, jwt_token

    except Exception as e:
        logger.error(f"Erro durante autenticação: {e}")
        raise

def _normalizar_cpf_cnpj(valor: Any) -> Tuple[str, int]:
    import math
    if valor is None:
        return "", 0
    if isinstance(valor, float):
        if math.isnan(valor):
            return "", 0
        if float(valor).is_integer():
            valor = int(valor)
    s = str(valor).strip()
    if "E+" in s.upper() or "E-" in s.upper():
        try:
            from decimal import Decimal
            s = format(Decimal(s), "f")
        except Exception:
            pass
    dig = re.sub(r"\D", "", s)
    if not dig:
        return "", 0
    if len(dig) <= 11:
        return dig.zfill(11), 1
    return dig.zfill(14), 2


def _buscar_pdf_b64(dados: Any) -> Optional[str]:
    if not dados:
        return None
    if isinstance(dados, dict):
        for k, v in dados.items():
            if isinstance(v, str) and v.startswith("JVBERi0"):  # assinatura típica base64 de PDF
                return v
            elif isinstance(v, (dict, list)):
                rec = _buscar_pdf_b64(v)
                if rec:
                    return rec
    elif isinstance(dados, list):
        for item in dados:
            rec = _buscar_pdf_b64(item)
            if rec:
                return rec
    return None


def baixar_pdf_guia(sistema_sn: str, cpf_cnpj: str, parcela_aaaamm: int, caminho_saida: Optional[Path] = None) -> bool:
    """
    Baixa e salva o PDF da guia de arrecadação do Integra Contador.

    Args:
        sistema_sn: Sistema (ex.: "PARCSN")
        cpf_cnpj: CPF ou CNPJ do contribuinte
        parcela_aaaamm: Parcela no formato AAAAMM (ex.: 202412)
        caminho_saida: Caminho para salvar o PDF (opcional, padrão: BASE_DIR / "PDFs" / "nome.pdf")

    Returns:
        True se sucesso, False se erro.
    """
    # CNPJ do contratante (deve estar no .env)
    cnpj_contratante = os.getenv("CNPJ_CONTRATANTE", "").strip()
    cnpj_contratante = re.sub(r"\D", "", cnpj_contratante).zfill(14)
    if not cnpj_contratante or len(cnpj_contratante) != 14:
        logger.error("Defina CNPJ_CONTRATANTE no .env (14 dígitos).")
        return False

    logger.info(f"Iniciando download de PDF para {sistema_sn} | {cpf_cnpj} | {parcela_aaaamm}")

    sistema_sn = sistema_sn.strip().upper()
    id_servico = ID_SERVICOS_EMISSAO.get(sistema_sn)
    if not id_servico:
        logger.error(f"idServico não configurado para {sistema_sn}")
        return False

    doc, tipo = _normalizar_cpf_cnpj(cpf_cnpj)
    if not doc:
        logger.error("CPF/CNPJ inválido após normalização.")
        return False

    try:
        access_token, jwt_token = obter_tokens()
    except Exception as e:
        logger.error(f"Erro na autenticação: {e}")
        return False

    payload = {
        "contratante": {"numero": cnpj_contratante, "tipo": 2},
        "autorPedidoDados": {"numero": cnpj_contratante, "tipo": 2},
        "contribuinte": {"numero": doc, "tipo": tipo},
        "pedidoDados": {
            "idSistema": sistema_sn,
            "idServico": id_servico,
            "versaoSistema": "1.0",
            "dados": json.dumps({"parcelaParaEmitir": int(parcela_aaaamm)}, ensure_ascii=False),
        },
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "jwt_token": jwt_token,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    try:
        resp = requests.post(
            f"{BASE_URL}/Emitir",
            headers=headers,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            timeout=60,
        )
    except Exception as e:
        logger.error(f"Erro na requisição: {e}")
        log_cobranca("Tentativa de Emissão Guia", valor=0.00, detalhes=f"Erro: {e}")
        return False

    try:
        raw = resp.json()
    except Exception:
        raw = resp.text

    if not resp.ok:
        msgs = []
        if isinstance(raw, dict):
            lst = raw.get("mensagens") or raw.get("erros") or []
            if isinstance(lst, list):
                for m in lst:
                    if isinstance(m, dict):
                        cod = m.get("codigo") or m.get("cod") or ""
                        desc = m.get("descricao") or m.get("texto") or m.get("mensagem") or ""
                        msgs.append(f"{cod}: {desc}".strip(": ").strip())
                    else:
                        msgs.append(str(m))
        if not msgs:
            msgs = [f"HTTP {resp.status_code} - {resp.reason}", (resp.text or "")[:500]]
        logger.error(f"Erro na API: {'; '.join(msgs)}")
        log_cobranca("Emissão Guia", valor=0.00, detalhes=f"Erro: {'; '.join(msgs)}")
        return False

    dados_str = raw.get("dados") if isinstance(raw, dict) else None
    if not dados_str:
        logger.error("Resposta sem campo 'dados'.")
        return False

    if isinstance(dados_str, str):
        try:
            dados = json.loads(dados_str)
        except Exception:
            dados = dados_str
    else:
        dados = dados_str

    pdf_b64 = _buscar_pdf_b64(dados)
    if not pdf_b64:
        logger.error("PDF base64 não localizado na resposta.")
        return False

    try:
        b64_clean = "".join(str(pdf_b64).split())
        pdf_bytes = base64.b64decode(b64_clean, validate=False)
        if not pdf_bytes.startswith(b"%PDF"):
            raise ValueError("Conteúdo decodificado não parece PDF (%PDF).")
    except Exception as e:
        logger.error(f"Falha ao decodificar PDF base64: {e}")
        return False

    if caminho_saida is None:
        nome_pdf = f"{sistema_sn}_{doc}_{parcela_aaaamm}.pdf".replace("/", "-").replace("\\", "-")
        caminho_saida = BASE_DIR / "PDFs" / nome_pdf

    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    with open(caminho_saida, "wb") as f:
        f.write(pdf_bytes)

    logger.info(f"PDF salvo em: {caminho_saida}")
    log_cobranca("Emissão Guia", valor=0.10, detalhes=f"Sistema: {sistema_sn}, CNPJ: {doc}, Parcela: {parcela_aaaamm}")

    return True
    """
    Autentica no Serpro (Integra Contador) e devolve (access_token, jwt_token).

    Requer no .env (na pasta mãe):
        SERPRO_CONSUMER_KEY
        SERPRO_CONSUMER_SECRET
        SERPRO_PFX_PATH
        SERPRO_PFX_PASSWORD
    """
    logger.info("Iniciando autenticação no Serpro.")

    consumer_key = os.getenv("SERPRO_CONSUMER_KEY")
    consumer_secret = os.getenv("SERPRO_CONSUMER_SECRET")
    pfx_path_raw = os.getenv("SERPRO_PFX_PATH")
    pfx_password = os.getenv("SERPRO_PFX_PASSWORD")

    if not all([consumer_key, consumer_secret, pfx_path_raw, pfx_password]):
        logger.error("Variáveis de ambiente faltando no .env.")
        raise RuntimeError(
            "Faltam variáveis no .env: "
            "SERPRO_CONSUMER_KEY, SERPRO_CONSUMER_SECRET, "
            "SERPRO_PFX_PATH, SERPRO_PFX_PASSWORD."
        )

    pfx_path = _limpar_caminho_pfx(pfx_path_raw)
    pfx_file = Path(pfx_path)

    if not pfx_file.exists():
        logger.error(f"Arquivo PFX não encontrado: {str(pfx_file)}")
        raise FileNotFoundError(
            "Caminho do certificado PFX não encontrado.\n"
            f"  SERPRO_PFX_PATH (bruto): {pfx_path_raw!r}\n"
            f"  Caminho após limpeza:   {str(pfx_file)!r}"
        )

    basic = base64.b64encode(f"{consumer_key}:{consumer_secret}".encode()).decode()

    url = "https://autenticacao.sapi.serpro.gov.br/authenticate"

    headers = {
        "Authorization": f"Basic {basic}",
        "role-type": "TERCEIROS",
        "content-type": "application/x-www-form-urlencoded",
    }

    data = {"grant_type": "client_credentials"}

    try:
        logger.info("Enviando requisição de autenticação para Serpro.")
        resp = pkcs12_post(
            url,
            data=data,
            headers=headers,
            pkcs12_filename=str(pfx_file),
            pkcs12_password=pfx_password,
            timeout=30  # Timeout de 30 segundos
        )
        resp.raise_for_status()
        j = resp.json()

        access_token = j["access_token"]
        jwt_token = j["jwt_token"]

        logger.info("Autenticação bem-sucedida. Tokens obtidos.")
        # Registrar possível cobrança pela autenticação
        log_cobranca("Autenticação Serpro (Integra Contador)", valor=0.00, detalhes="Obtenção de tokens OAuth2")

        return access_token, jwt_token

    except Exception as e:
        logger.error(f"Erro durante autenticação: {e}")
        raise


def processar_lista_parcelamentos(caminho_excel: Optional[Path] = None) -> None:
    """
    Lê a lista de parcelamentos do Excel e identifica quais precisam de PDFs.
    Por enquanto, apenas analisa e loga; não baixa ainda.
    """
    if caminho_excel is None:
        caminho_excel = BASE_DIR / "LISTA PARCELAMENTOS.xlsx"

    if not caminho_excel.exists():
        logger.error(f"Arquivo Excel não encontrado: {caminho_excel}")
        return

    logger.info(f"Lendo lista de parcelamentos: {caminho_excel}")

    try:
        df = pd.read_excel(caminho_excel)
    except Exception as e:
        logger.error(f"Erro ao ler Excel: {e}")
        return

    logger.info(f"Colunas encontradas: {list(df.columns)}")
    logger.info(f"Total de linhas: {len(df)}")

    # Assumir colunas baseadas no arquivo real
    col_cpf = "CPF/CNPJ"
    col_parcelamento = "PARCELAMENTO"
    col_enviado = "ENVIADO"

    if col_cpf not in df.columns or col_parcelamento not in df.columns:
        logger.error(f"Colunas esperadas '{col_cpf}' ou '{col_parcelamento}' não encontradas.")
        return

    # Filtrar linhas com parcelamento preenchido e não enviado
    df_filtrado = df[df[col_parcelamento].notna() & df[col_enviado].isna()]

    logger.info(f"Linhas com parcelamento pendente: {len(df_filtrado)}")

    for idx, row in df_filtrado.iterrows():
        cpf_cnpj = str(row[col_cpf]).strip()
        parcelamento_desc = str(row[col_parcelamento]).strip()

        logger.info(f"Linha {idx+1}: CPF/CNPJ={cpf_cnpj}, Parcelamento='{parcelamento_desc}'")

        # Tentar extrair sistema e parcela da descrição
        sistema, parcela = extrair_sistema_parcela(parcelamento_desc)
        if sistema and parcela:
            logger.info(f"  -> Sistema: {sistema}, Parcela: {parcela}")
            # Baixar o PDF
            sucesso = baixar_pdf_guia(sistema, cpf_cnpj, parcela)
            if sucesso:
                logger.info(f"  -> PDF baixado com sucesso para linha {idx+1}")
            else:
                logger.error(f"  -> Falha ao baixar PDF para linha {idx+1}")
        else:
            logger.warning(f"  -> Não conseguiu extrair sistema/parcela de '{parcelamento_desc}' - pulando download")

    logger.info("Análise concluída. Ajuste o Excel ou a lógica de extração conforme necessário.")


def extrair_sistema_parcela(descricao: str) -> Tuple[Optional[str], Optional[int]]:
    """
    Extrai sistema e parcela da descrição do parcelamento.
    Exemplo: "PGFN N° 001.460.693" -> ("PGFN", 1460693)
    """
    descricao = descricao.upper()

    if "PGFN" in descricao:
        # Procurar número após PGFN
        import re
        match = re.search(r'PGFN.*?(\d{1,3}(?:\.\d{3})*\d{3})', descricao)
        if match:
            numero = match.group(1).replace('.', '')
            return "PGFN", int(numero)
    elif "SIMPLIFICADO" in descricao:
        return "SIMPLIFICADO", None  # Parcela pode ser extraída de outro lugar
    elif "PREVIDENCIARIO" in descricao:
        return "PREVIDENCIARIO", None

    return None, None


if __name__ == "__main__":
    import sys
    if len(sys.argv) == 2 and sys.argv[1] == "processar":
        # Modo processar lista: python auth_integra.py processar
        print("Processando lista de parcelamentos...")
        processar_lista_parcelamentos()
        print("Processamento concluído.")
    elif len(sys.argv) == 4:
        # Modo download: python auth_integra.py PARCSN 12345678000123 202412
        sistema, cpf_cnpj, parcela = sys.argv[1], sys.argv[2], int(sys.argv[3])
        print(f"Baixando PDF para {sistema} | {cpf_cnpj} | {parcela}...")
        sucesso = baixar_pdf_guia(sistema, cpf_cnpj, parcela)
        if sucesso:
            print("PDF baixado com sucesso!")
        else:
            print("Erro ao baixar PDF.")
            sys.exit(1)
    else:
        # Modo teste de autenticação
        try:
            access_token, jwt_token = obter_tokens()
            print("Autenticação bem-sucedida!")
            print(f"Access Token: {access_token}")
            print(f"JWT Token: {jwt_token}")
        except Exception as e:
            print(f"Erro: {e}")
            sys.exit(1)
