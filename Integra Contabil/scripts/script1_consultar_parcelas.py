# -*- coding: utf-8 -*-
"""
script1_consultar_parcelas.py

Consulta de parcelas disponíveis para emissão de guias no Integra Contador
(Simples Nacional) - MODO COMPLETO

Entrada:
  - LISTA PARCELAMENTOS.xlsx (na pasta mãe)
  - auth_integra.py (na pasta mãe)
  - .env (na pasta mãe)

Saídas:
  - arquivos/api_teste/controle_uso_api.xlsx
  - arquivos/api_teste/parcelas_encontradas.xlsx
  - arquivos/api_teste/parcelas_todas_situacoes.xlsx
  - arquivos/api_teste/sem_procuracao.xlsx
  - arquivos/api_teste/sem_parcelamento_ativo.xlsx

Regras:
  - Lê a planilha inteira como TEXTO (não corta zeros)
  - Remove separadores em branco e cabeçalhos repetidos no meio
  - Considera SOMENTE Simples Nacional (exclui "SIMPLIFIC...")
"""

import os
import sys
import json
import re
import math
import datetime as dt
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
from dotenv import load_dotenv

# =========================
# LOCALIZAÇÃO DA PASTA MÃE
# =========================

NOME_ARQ_LISTA = "LISTA PARCELAMENTOS.xlsx"
NOME_AUTH = "auth_integra.py"


def encontrar_pasta_mae(
    nome_arquivo_lista: str = NOME_ARQ_LISTA,
    nome_auth: str = NOME_AUTH,
    max_niveis: int = 6,
) -> Path:
    pasta = Path(__file__).resolve().parent
    for _ in range(max_niveis):
        if (pasta / nome_arquivo_lista).exists() and (pasta / nome_auth).exists():
            return pasta
        pasta = pasta.parent
    raise FileNotFoundError(
        f'Não encontrei a pasta mãe contendo "{nome_arquivo_lista}" e "{nome_auth}" '
        f"nos {max_niveis} níveis acima."
    )


PASTA_MAE = encontrar_pasta_mae()

if str(PASTA_MAE) not in sys.path:
    sys.path.insert(0, str(PASTA_MAE))

from auth_integra import obter_tokens  # type: ignore

load_dotenv(PASTA_MAE / ".env")

PASTA_ARQUIVOS_API = PASTA_MAE / "arquivos" / "api_teste"
PASTA_ARQUIVOS_API.mkdir(parents=True, exist_ok=True)

ARQ_ENTRADA = PASTA_MAE / NOME_ARQ_LISTA
ARQ_CONTROLE = PASTA_ARQUIVOS_API / "controle_uso_api.xlsx"
ARQ_PARCELAS_ABERTO = PASTA_ARQUIVOS_API / "parcelas_encontradas.xlsx"
ARQ_PARCELAS_TODAS = PASTA_ARQUIVOS_API / "parcelas_todas_situacoes.xlsx"
ARQ_SEM_PROC = PASTA_ARQUIVOS_API / "sem_procuracao.xlsx"
ARQ_SEM_PARC_ATIVO = PASTA_ARQUIVOS_API / "sem_parcelamento_ativo.xlsx"

# =========================
# CONFIG INTEGRA
# =========================

BASE_URL = "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1"

CNPJ_CONTRATANTE = os.getenv("CNPJ_CONTRATANTE", "").strip()
CNPJ_CONTRATANTE = re.sub(r"\D", "", CNPJ_CONTRATANTE).zfill(14)
if not CNPJ_CONTRATANTE or len(CNPJ_CONTRATANTE) != 14:
    raise RuntimeError("Defina CNPJ_CONTRATANTE no .env (14 dígitos).")

ID_SERVICOS_CONSULTA: Dict[str, str] = {
    "PARCSN": "PARCELASPARAGERAR162",
    "PARCSN-ESP": "PARCELASPARAGERAR172",
    "PERTSN": "PARCELASPARAGERAR182",
    "RELPSN": "PARCELASPARAGERAR192",
}

# =========================
# NORMALIZAÇÕES
# =========================


def normalizar_texto(x: Any) -> str:
    if x is None:
        return ""
    s = str(x).upper()
    for ac, sem in (
        ("Á", "A"), ("À", "A"), ("Ã", "A"), ("Â", "A"),
        ("É", "E"), ("Ê", "E"),
        ("Í", "I"),
        ("Ó", "O"), ("Õ", "O"), ("Ô", "O"),
        ("Ú", "U"), ("Ü", "U"),
        ("Ç", "C"),
    ):
        s = s.replace(ac, sem)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalizar_cpf_cnpj(valor: Any) -> Tuple[str, int]:
    """
    Retorna (doc_digits, tipo_doc)
      - CPF:  11 dígitos (zfill), tipo=1
      - CNPJ: 14 dígitos (zfill), tipo=2
    """
    if valor is None:
        return "", 0

    if isinstance(valor, float):
        if math.isnan(valor):
            return "", 0
        if float(valor).is_integer():
            valor = int(valor)

    s = str(valor).strip()
    # trata notação científica se vier como string
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


def detectar_sistema_sn(parcelamento: Any) -> Optional[str]:
    """
    SOMENTE Simples Nacional (EXCLUI qualquer "SIMPLIFIC...").
    Aceita variantes:
      - SIMPLES
      - SIMPLES NACIONAL
      - SIMPLES NACIONAL ESPECIAL
      - SIMPLES PERT
      - SIMPLES RELP
    """
    t = normalizar_texto(parcelamento)
    if not t:
        return None

    if "SIMPLIFIC" in t:
        return None

    if "SIMPLES" not in t:
        return None

    if "RELP" in t:
        return "RELPSN"
    if "PERT" in t:
        return "PERTSN"
    if "ESPECIAL" in t:
        return "PARCSN-ESP"
    return "PARCSN"


def formatar_mensagens_humanas(mensagens: List[str]) -> str:
    limpas: List[str] = []
    for m in mensagens or []:
        if not m:
            continue
        s = str(m).replace("\n", " ").replace("\r", " ").strip()
        s = re.sub(r"\s+", " ", s)
        limpas.append(s)
    return " ; ".join(limpas)


# =========================
# LEITURA ROBUSTA DO EXCEL
# =========================


def _achar_coluna(df: pd.DataFrame, contains_upper: str) -> str:
    cand = [c for c in df.columns if contains_upper in normalizar_texto(c)]
    if not cand:
        raise KeyError(f'Não encontrei coluna contendo "{contains_upper}" em {list(df.columns)}')
    # se tiver mais de uma, pega a última (geralmente a mais recente)
    return cand[-1]


def ler_lista_parcelamentos() -> Tuple[pd.DataFrame, str, str, str]:
    if not ARQ_ENTRADA.exists():
        raise FileNotFoundError(f'Arquivo "{ARQ_ENTRADA}" não encontrado.')

    xls = pd.ExcelFile(ARQ_ENTRADA, engine="openpyxl")
    dfs: List[pd.DataFrame] = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(ARQ_ENTRADA, sheet_name=sheet, dtype=str, na_filter=False, engine="openpyxl")
        if df.empty:
            continue
        df["__SHEET__"] = sheet
        dfs.append(df)

    if not dfs:
        raise RuntimeError("Nenhuma aba com dados encontrada no Excel.")

    df = pd.concat(dfs, ignore_index=True)

    col_emp = _achar_coluna(df, "EMPRESA")
    col_doc = _achar_coluna(df, "CPF")
    col_parc = _achar_coluna(df, "PARCELAMENTO")

    # trims
    for c in (col_emp, col_doc, col_parc):
        df[c] = df[c].astype(str).str.strip()

    # remove separadores totalmente vazios (linha preta)
    sep = (df[col_emp] == "") & (df[col_doc] == "") & (df[col_parc] == "")
    df = df[~sep].copy()

    # remove cabeçalho repetido no meio (linha que vira "dados")
    cod_col = None
    for c in df.columns:
        if normalizar_texto(c) in ("CÓD.", "COD.", "COD", "CÓD"):
            cod_col = c
            break

    header_rep = (
        df[col_doc].str.upper().isin(["CPF/CNPJ", "CPF", "CNPJ"])
        | (df[col_parc].str.upper() == "PARCELAMENTO")
        | df[col_emp].str.upper().str.startswith("EMPRESA")
    )
    if cod_col:
        header_rep = header_rep | (df[cod_col].astype(str).str.upper().isin(["CÓD.", "COD.", "COD", "CÓD"]))

    df = df[~header_rep].copy()

    return df, col_emp, col_doc, col_parc


# =========================
# CONTROLE / CACHE
# =========================


def carregar_controle() -> pd.DataFrame:
    cols = [
        "TIPO_OPERACAO",
        "SISTEMA_SN",
        "CPF_CNPJ",
        "ID_PARCELAMENTO",
        "ID_PARCELA_API",
        "NUMERO_PARCELA",
        "DATA_HORA",
        "RESULTADO",
        "DETALHE",
        "DADOS_BRUTOS_JSON",
    ]
    if ARQ_CONTROLE.exists():
        dfc = pd.read_excel(ARQ_CONTROLE, dtype=str, na_filter=False, engine="openpyxl")
        for c in cols:
            if c not in dfc.columns:
                dfc[c] = ""
        return dfc[cols].copy()
    return pd.DataFrame(columns=cols)


def montar_cache_consultas(dfc: pd.DataFrame) -> Dict[Tuple[str, str], Any]:
    cache: Dict[Tuple[str, str], Any] = {}
    if dfc.empty:
        return cache

    mask = (dfc["TIPO_OPERACAO"].astype(str).str.upper() == "CONSULTA_PARCELAS_API") & (
        dfc["RESULTADO"].astype(str).str.upper() == "SUCESSO"
    )
    df_ok = dfc[mask].copy()

    for _, r in df_ok.iterrows():
        sistema = str(r.get("SISTEMA_SN", "")).strip().upper()
        doc = normalizar_cpf_cnpj(r.get("CPF_CNPJ", ""))[0]
        dados_json = r.get("DADOS_BRUTOS_JSON", "")
        if not sistema or not doc or not str(dados_json).strip():
            continue
        try:
            cache[(sistema, doc)] = json.loads(dados_json)
        except Exception:
            continue

    return cache


# =========================
# API - CONSULTAR
# =========================


def consultar_parcelas_disponiveis(sistema_sn: str, cpf_cnpj: str) -> Dict[str, Any]:
    sistema_sn = sistema_sn.strip().upper()
    id_servico = ID_SERVICOS_CONSULTA.get(sistema_sn)
    if not id_servico:
        return {"ok": False, "mensagens": [f"idServico não configurado para {sistema_sn}"], "dados_raw": None}

    doc, tipo = normalizar_cpf_cnpj(cpf_cnpj)
    if not doc:
        return {"ok": False, "mensagens": ["CPF/CNPJ inválido após normalização."], "dados_raw": None}

    access_token, jwt_token = obter_tokens()

    payload = {
        "contratante": {"numero": CNPJ_CONTRATANTE, "tipo": 2},
        "autorPedidoDados": {"numero": CNPJ_CONTRATANTE, "tipo": 2},
        "contribuinte": {"numero": doc, "tipo": tipo},
        "pedidoDados": {
            "idSistema": sistema_sn,
            "idServico": id_servico,
            "versaoSistema": "1.0",
            "dados": "",  # conforme regra do Integra p/ esses serviços
        },
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "jwt_token": jwt_token,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    resp = requests.post(
        f"{BASE_URL}/Consultar",
        headers=headers,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        timeout=60,
    )

    try:
        raw = resp.json()
    except Exception:
        raw = resp.text

    if not resp.ok:
        msgs: List[str] = []
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
        return {"ok": False, "mensagens": msgs, "dados_raw": raw}

    dados_str = raw.get("dados") if isinstance(raw, dict) else None
    mensagens_raw = raw.get("mensagens", []) if isinstance(raw, dict) else []
    msgs: List[str] = []
    if isinstance(mensagens_raw, list):
        for m in mensagens_raw:
            if isinstance(m, dict):
                cod = m.get("codigo") or m.get("cod") or ""
                desc = m.get("descricao") or m.get("texto") or m.get("mensagem") or ""
                msgs.append(f"{cod}: {desc}".strip(": ").strip())
            else:
                msgs.append(str(m))

    if not dados_str:
        return {"ok": True, "mensagens": msgs, "dados_raw": {}}

    if isinstance(dados_str, str):
        try:
            dados = json.loads(dados_str)
        except Exception:
            dados = dados_str
    else:
        dados = dados_str

    return {"ok": True, "mensagens": msgs, "dados_raw": dados}


def extrair_lista_parcelas(dados_raw: Any) -> List[Dict[str, Any]]:
    if dados_raw is None:
        return []
    if isinstance(dados_raw, list):
        return [p for p in dados_raw if isinstance(p, dict)]
    if isinstance(dados_raw, dict):
        for k in ("parcelas", "parcelasDisponiveis", "listaParcelas", "listaParcelasParaGeracao"):
            v = dados_raw.get(k)
            if isinstance(v, list):
                return [p for p in v if isinstance(p, dict)]
        # fallback: procura primeira lista de dicts dentro do dict
        for v in dados_raw.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return v
    return []


def extrair_campo(p: Dict[str, Any], chaves: Tuple[str, ...]) -> Any:
    for k in chaves:
        if k in p:
            return p[k]
    return None


def situacao_eh_em_aberto(situacao: Any) -> bool:
    if situacao is None:
        return True  # quando API não manda, considerar disponível p/ emitir
    s = normalizar_texto(situacao)
    if not s:
        return True
    if any(x in s for x in ("PAGO", "QUITADO", "LIQUIDADO", "CANCELADO")):
        return False
    if any(x in s for x in ("ABERTO", "EM ABERTO", "DISPONIVEL", "PENDENTE", "A VENCER", "NAO PAGO")):
        return True
    return False


def classificar_situacao_controle(det: str) -> str:
    s = normalizar_texto(det)
    if "ACESSONEGADO-ICGERENCIADOR-022" in s or ("ACESSO NEGADO" in s and "PROCURA" in s):
        return "SEM_PROCURACAO"
    if "NAO HA PARCELAMENTO ATIVO" in s or "NAO HA PARCELAMENTO RELP-SN ATIVO" in s:
        return "SEM_PARCELAMENTO_ATIVO"
    return ""


# =========================
# MAIN
# =========================


def main() -> None:
    df_raw, col_emp, col_doc, col_parc = ler_lista_parcelamentos()

    # Normaliza documento e detecta sistema
    df_raw["CPF_CNPJ"] = df_raw[col_doc].apply(lambda v: normalizar_cpf_cnpj(v)[0])
    df_raw["SISTEMA_SN"] = df_raw[col_parc].apply(detectar_sistema_sn)

    # Filtra só SN
    df_sn = df_raw[df_raw["SISTEMA_SN"].notna()].copy()
    df_sn = df_sn[df_sn["CPF_CNPJ"].astype(str).str.len() > 0].copy()

    print(f"Linhas totais lidas: {len(df_raw)}")
    print(f"Linhas Simples Nacional detectadas: {len(df_sn)}")

    if df_sn.empty:
        print("Nenhuma linha de Simples Nacional encontrada.")
        return

    # Controle / cache
    dfc = carregar_controle()
    cache = montar_cache_consultas(dfc)

    # Chaves únicas
    chaves = df_sn[["SISTEMA_SN", "CPF_CNPJ"]].drop_duplicates().copy()
    chaves["SISTEMA_SN"] = chaves["SISTEMA_SN"].astype(str).str.upper().str.strip()
    chaves["CPF_CNPJ"] = chaves["CPF_CNPJ"].astype(str).str.strip()

    print("\nPares únicos (SISTEMA_SN, CPF_CNPJ) a processar:")
    for _, r in chaves.iterrows():
        print(f"  - {r['SISTEMA_SN']} | {r['CPF_CNPJ']}")

    novos_logs: List[Dict[str, Any]] = []

    for _, r in chaves.iterrows():
        sistema = str(r["SISTEMA_SN"]).strip().upper()
        doc = str(r["CPF_CNPJ"]).strip()
        chave = (sistema, doc)

        if chave in cache:
            continue

        ret = consultar_parcelas_disponiveis(sistema, doc)
        ok = bool(ret.get("ok"))
        msgs = ret.get("mensagens", [])
        dados_raw = ret.get("dados_raw")

        if ok:
            cache[chave] = dados_raw

        novos_logs.append(
            {
                "TIPO_OPERACAO": "CONSULTA_PARCELAS_API",
                "SISTEMA_SN": sistema,
                "CPF_CNPJ": doc,
                "ID_PARCELAMENTO": "",
                "ID_PARCELA_API": "",
                "NUMERO_PARCELA": "",
                "DATA_HORA": dt.datetime.now().isoformat(timespec="seconds"),
                "RESULTADO": "SUCESSO" if ok else "ERRO",
                "DETALHE": formatar_mensagens_humanas(msgs),
                "DADOS_BRUTOS_JSON": json.dumps(dados_raw, ensure_ascii=False) if dados_raw is not None else "",
            }
        )

    if novos_logs:
        dfc = pd.concat([dfc, pd.DataFrame(novos_logs)], ignore_index=True)
        dfc.to_excel(ARQ_CONTROLE, index=False)

    # Resumos
    dfc2 = dfc.copy()
    if "TIPO_SITUACAO" not in dfc2.columns:
        dfc2["TIPO_SITUACAO"] = ""
    mask_cons = dfc2["TIPO_OPERACAO"].astype(str).str.upper() == "CONSULTA_PARCELAS_API"
    dfc2.loc[mask_cons, "TIPO_SITUACAO"] = dfc2.loc[mask_cons, "DETALHE"].astype(str).apply(classificar_situacao_controle)

    df_sem_proc = dfc2[mask_cons & (dfc2["TIPO_SITUACAO"] == "SEM_PROCURACAO")][["CPF_CNPJ", "SISTEMA_SN", "DETALHE", "DATA_HORA"]].drop_duplicates()
    if not df_sem_proc.empty:
        df_sem_proc.sort_values(by=["CPF_CNPJ", "SISTEMA_SN"]).to_excel(ARQ_SEM_PROC, index=False)

    df_sem_parc = dfc2[mask_cons & (dfc2["TIPO_SITUACAO"] == "SEM_PARCELAMENTO_ATIVO")][["CPF_CNPJ", "SISTEMA_SN", "DETALHE", "DATA_HORA"]].drop_duplicates()
    if not df_sem_parc.empty:
        df_sem_parc.sort_values(by=["CPF_CNPJ", "SISTEMA_SN"]).to_excel(ARQ_SEM_PARC_ATIVO, index=False)

    # Monta parcelas
    linhas_todas: List[Dict[str, Any]] = []
    linhas_aberto: List[Dict[str, Any]] = []

    for _, row in df_sn.iterrows():
        empresa = row[col_emp]
        doc = row["CPF_CNPJ"]
        sistema = str(row["SISTEMA_SN"]).strip().upper()
        desc_parc = row[col_parc]

        dados_raw = cache.get((sistema, doc))
        if dados_raw is None:
            continue

        parcelas = extrair_lista_parcelas(dados_raw)

        for p in parcelas:
            numero_parcela = extrair_campo(p, ("numeroParcela", "nrParcela", "numParcela", "parcela"))
            situacao = extrair_campo(p, ("situacao", "situacaoParcela", "status"))
            vencimento = extrair_campo(p, ("dataVencimento", "dtVencimento", "vencimento"))
            valor = extrair_campo(p, ("valorParcela", "vlParcela", "valor"))
            id_parcela_api = extrair_campo(p, ("idParcela", "identificadorParcela", "id"))

            linhas_todas.append(
                {
                    "EMPRESA": empresa,
                    "CPF_CNPJ": doc,
                    "SISTEMA_SN": sistema,
                    "DESCR_PARCELAMENTO_ORIG": desc_parc,
                    "ID_PARCELA_API": id_parcela_api,
                    "NUMERO_PARCELA": numero_parcela,
                    "DATA_VENCIMENTO": vencimento,
                    "VALOR_PARCELA": valor,
                    "SITUACAO_API": situacao,
                }
            )

            if not situacao_eh_em_aberto(situacao):
                continue

            linhas_aberto.append(
                {
                    "EMPRESA": empresa,
                    "CPF_CNPJ": doc,
                    "SISTEMA_SN": sistema,
                    "DESCR_PARCELAMENTO_ORIG": desc_parc,
                    "ID_PARCELA_API": id_parcela_api,
                    "NUMERO_PARCELA": numero_parcela,
                    "DATA_VENCIMENTO": vencimento,
                    "VALOR_PARCELA": valor,
                    "SITUACAO_API": situacao,
                    "EMITIDO": False,
                    "DATA_EMISSAO": "",
                    "CAMINHO_ARQUIVO": "",
                }
            )

    if linhas_todas:
        df_all = pd.DataFrame(linhas_todas)
        df_all.to_excel(ARQ_PARCELAS_TODAS, index=False)
    else:
        print("Nenhuma parcela retornada pela API (todas as situações).")

    if linhas_aberto:
        df_open = pd.DataFrame(linhas_aberto)
        df_open.to_excel(ARQ_PARCELAS_ABERTO, index=False)
    else:
        print("Nenhuma parcela em aberto encontrada para emissão.")

    print("\n=== RESUMO SCRIPT 1 ===")
    print(f"Controle: {ARQ_CONTROLE}")
    print(f"Parcelas (todas): {ARQ_PARCELAS_TODAS}")
    print(f"Parcelas (em aberto): {ARQ_PARCELAS_ABERTO}")
    print(f"Sem procuração: {ARQ_SEM_PROC if ARQ_SEM_PROC.exists() else '(nenhum)'}")
    print(f"Sem parcelamento ativo: {ARQ_SEM_PARC_ATIVO if ARQ_SEM_PARC_ATIVO.exists() else '(nenhum)'}")


if __name__ == "__main__":
    main()
