"""
script1_consultar_parcelas_teste.py

>> SCRIPT SOMENTE PARA TESTE <<
NÃO chama API, NÃO usa certificado, NÃO gasta nada.

Entrada:
    LISTA PARCELAMENTOS.xlsx  (na PASTA MÃE)

Saída (na PASTA_MAE/arquivos/teste):
    controle_uso_api_teste.xlsx
    parcelas_para_emissao_teste.xlsx
"""

import os
import json
import datetime as dt
from typing import Optional, Dict, Any, List, Tuple

import pandas as pd
from pathlib import Path

# =========================
# LOCALIZAÇÃO DA PASTA MÃE
# =========================

NOME_ARQ_LISTA = "LISTA PARCELAMENTOS.xlsx"

def encontrar_pasta_mae_com_lista(nome_arquivo: str = NOME_ARQ_LISTA, max_niveis: int = 5) -> Path:
    """
    Sobe a árvore de pastas a partir da pasta do script até encontrar
    o arquivo da lista de parcelas.
    """
    pasta = Path(__file__).resolve().parent
    for _ in range(max_niveis):
        if (pasta / nome_arquivo).exists():
            return pasta
        pasta = pasta.parent
    # fallback: se não achar, usa a pasta do próprio script
    return Path(__file__).resolve().parent

PASTA_MAE = encontrar_pasta_mae_com_lista()

# Subpasta organizada para saídas de teste
PASTA_ARQUIVOS_TESTE = PASTA_MAE / "arquivos" / "teste"
PASTA_ARQUIVOS_TESTE.mkdir(parents=True, exist_ok=True)

ARQ_ENTRADA = PASTA_MAE / NOME_ARQ_LISTA
ARQ_VALIDACAO = PASTA_ARQUIVOS_TESTE / "controle_uso_api_teste.xlsx"
ARQ_PARA_EMISSAO = PASTA_ARQUIVOS_TESTE / "parcelas_para_emissao_teste.xlsx"


# =========================
# DETECTAR SISTEMA (PARCSN / PERTSN / RELPSN / PARCSN-ESP)
# =========================

def detectar_sistema_sn(descricao_parcelamento: Optional[str]) -> Optional[str]:
    if not descricao_parcelamento or pd.isna(descricao_parcelamento):
        return None

    texto = str(descricao_parcelamento).upper()

    if "SIMPLES" not in texto:
        return None

    if "RELP" in texto:
        return "RELPSN"
    if "PERT" in texto:
        return "PERTSN"
    if "ESPECIAL" in texto:
        return "PARCSN-ESP"

    return "PARCSN"


# =========================
# CONTROLE / VALIDAÇÃO (SÓ TESTE)
# =========================

def carregar_controle_teste() -> pd.DataFrame:
    if ARQ_VALIDACAO.exists():
        return pd.read_excel(ARQ_VALIDACAO)

    colunas = [
        "TIPO_OPERACAO",
        "SISTEMA_SN",
        "CPF_CNPJ",
        "ID_PARCELAMENTO",
        "ID_PARCELA_API",
        "DATA_HORA",
        "RESULTADO",
        "DETALHE",
        "DADOS_BRUTOS_JSON"
    ]
    return pd.DataFrame(columns=colunas)


def salvar_controle_teste(df_controle: pd.DataFrame) -> None:
    df_controle.to_excel(ARQ_VALIDACAO, index=False)


# =========================
# "CONSULTA" FALSA (SEM API)
# =========================

def consultar_parcelas_disponiveis_teste(
    sistema_sn: str, cpf_cnpj: str
) -> Dict[str, Any]:
    base_id = f"TEST-{sistema_sn}-{cpf_cnpj}"

    dados_fake = {
        "parcelas": [
            {
                "numeroParcela": 1,
                "situacao": "EM ABERTO",
                "dataVencimento": "2025-01-10",
                "valorParcela": 123.45,
                "idParcela": f"{base_id}-1"
            },
            {
                "numeroParcela": 2,
                "situacao": "PAGA",
                "dataVencimento": "2025-02-10",
                "valorParcela": 150.00,
                "idParcela": f"{base_id}-2"
            }
        ]
    }

    mensagens = [f"TESTE: dados fictícios gerados para {cpf_cnpj} / {sistema_sn}"]

    return {
        "ok": True,
        "mensagens": mensagens,
        "lista_parcelas": dados_fake["parcelas"],
        "dados_raw": dados_fake
    }


# =========================
# AJUDANTES PARA CAMPOS DA PARCELA
# =========================

def extrair_campo_parcela(
    parcela: Dict[str, Any],
    possiveis_nomes: Tuple[str, ...]
) -> Any:
    for nome in possiveis_nomes:
        if nome in parcela:
            return parcela[nome]
    return None


def situacao_eh_em_aberto(situacao: Optional[str]) -> bool:
    if not situacao:
        return False
    s = situacao.upper().replace("Í", "I").replace("Ã", "A").strip()
    return s in {
        "EM ABERTO",
        "EMABERTO",
        "ABERTA",
        "ABERTO",
        "DISPONIVEL",
        "DISPONIVEL PARA PAGAMENTO",
        "PENDENTE",
    }


# =========================
# SCRIPT PRINCIPAL (SÓ TESTE)
# =========================

def main() -> None:
    # 1) Carrega Excel original da PASTA_MAE
    if not ARQ_ENTRADA.exists():
        raise FileNotFoundError(
            f'Arquivo "{ARQ_ENTRADA}" não encontrado. '
            f'Certifique-se de que "{NOME_ARQ_LISTA}" está na pasta mãe: {PASTA_MAE}'
        )

    df = pd.read_excel(ARQ_ENTRADA)

    col_empresa = "EMPRESA / PESSOA FISICA 07/2022"
    col_doc = "CPF/CNPJ"
    col_parcelamento = "PARCELAMENTO"

    for col in (col_empresa, col_doc, col_parcelamento):
        if col not in df.columns:
            raise KeyError(
                f'A coluna "{col}" não foi encontrada em {ARQ_ENTRADA.name}. '
                f'Confirme que o layout não foi alterado.'
            )

    df[col_doc] = (
        df[col_doc]
        .astype(str)
        .str.replace(r"\D", "", regex=True)
        .str.strip()
    )
    df[col_parcelamento] = df[col_parcelamento].astype(str).str.strip()

    # 2) Detecta sistema SN
    df["SISTEMA_SN"] = df[col_parcelamento].apply(detectar_sistema_sn)

    df_sn = df[df["SISTEMA_SN"].notna()].copy()
    if df_sn.empty:
        print("Nenhuma linha de Simples Nacional encontrada.")
        return

    # 3) Controle / cache (somente arquivos de teste)
    df_controle = carregar_controle_teste()
    registros_controle_novos: List[Dict[str, Any]] = []
    cache_parcelas_por_chave: Dict[Tuple[str, str], Dict[str, Any]] = {}

    chaves = (
        df_sn[["SISTEMA_SN", col_doc]]
        .drop_duplicates()
        .rename(columns={col_doc: "CPF_CNPJ"})
    )

    for _, linha in chaves.iterrows():
        sistema_sn = linha["SISTEMA_SN"]
        cpf_cnpj = linha["CPF_CNPJ"]

        retorno = consultar_parcelas_disponiveis_teste(sistema_sn, cpf_cnpj)
        ok = retorno["ok"]
        dados_raw = retorno["dados_raw"]
        mensagens = retorno["mensagens"]

        if ok and dados_raw is not None:
            cache_parcelas_por_chave[(sistema_sn, cpf_cnpj)] = dados_raw

        registros_controle_novos.append({
            "TIPO_OPERACAO": "CONSULTA_PARCELAS",
            "SISTEMA_SN": sistema_sn,
            "CPF_CNPJ": cpf_cnpj,
            "ID_PARCELAMENTO": None,
            "ID_PARCELA_API": None,
            "DATA_HORA": dt.datetime.now().isoformat(timespec="seconds"),
            "RESULTADO": "SUCESSO" if ok else "ERRO",
            "DETALHE": json.dumps(mensagens, ensure_ascii=False),
            "DADOS_BRUTOS_JSON": (
                json.dumps(dados_raw, ensure_ascii=False)
                if dados_raw is not None
                else None
            ),
        })

    if registros_controle_novos:
        df_controle = pd.concat(
            [df_controle, pd.DataFrame(registros_controle_novos)],
            ignore_index=True
        )
        salvar_controle_teste(df_controle)

    # 4) Monta planilha de "parcelas para emissão" (teste)
    linhas_emissao: List[Dict[str, Any]] = []

    for _, row in df_sn.iterrows():
        sistema_sn = row["SISTEMA_SN"]
        cpf_cnpj = row[col_doc]
        empresa = row[col_empresa]
        desc_parcelamento = row[col_parcelamento]

        dados_raw = cache_parcelas_por_chave.get((sistema_sn, cpf_cnpj))
        if not dados_raw:
            continue

        if isinstance(dados_raw, dict):
            lista_parcelas = dados_raw.get("parcelas", [])
        elif isinstance(dados_raw, list):
            lista_parcelas = dados_raw
        else:
            lista_parcelas = []

        for p in lista_parcelas:
            if not isinstance(p, dict):
                continue

            numero_parcela = extrair_campo_parcela(
                p,
                ("numeroParcela", "nrParcela", "numParcela")
            )
            situacao = extrair_campo_parcela(
                p,
                ("situacao", "situacaoParcela", "status")
            )
            vencimento = extrair_campo_parcela(
                p,
                ("dataVencimento", "dtVencimento", "vencimento")
            )
            valor = extrair_campo_parcela(
                p,
                ("valorParcela", "vlParcela", "valor")
            )
            id_parcela_api = extrair_campo_parcela(
                p,
                ("idParcela", "identificadorParcela", "id")
            )

            if not situacao_eh_em_aberto(situacao):
                continue

            linhas_emissao.append({
                "EMPRESA": empresa,
                "CPF_CNPJ": cpf_cnpj,
                "SISTEMA_SN": sistema_sn,
                "DESCR_PARCELAMENTO_ORIG": desc_parcelamento,
                "ID_PARCELA_API": id_parcela_api,
                "NUMERO_PARCELA": numero_parcela,
                "DATA_VENCIMENTO": vencimento,
                "VALOR_PARCELA": valor,
                "SITUACAO_API": situacao,
                "EMITIDO": False,
                "DATA_EMISSAO": None,
                "CAMINHO_ARQUIVO": None,
            })

    if not linhas_emissao:
        print("Nenhuma parcela elegível (em aberto) encontrada – mesmo no teste.")
        return

    df_emissao = pd.DataFrame(linhas_emissao)
    df_emissao.to_excel(ARQ_PARA_EMISSAO, index=False)

    print("Concluído (MODO TESTE, sem API).")
    print(f"- PASTA MÃE: {PASTA_MAE}")
    print(f"- Arquivos gerados em: {PASTA_ARQUIVOS_TESTE}")
    print(f"- Controle (teste): {ARQ_VALIDACAO.name}")
    print(f"- Parcelas p/ emissão (teste): {ARQ_PARA_EMISSAO.name}")


if __name__ == "__main__":
    main()
