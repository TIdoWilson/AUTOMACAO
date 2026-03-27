# -*- coding: utf-8 -*-
"""
script2_emissao_documentos.py

Emissão de DAS (Integra Parcelamento - Simples Nacional)

Entrada:
  - arquivos/api_teste/parcelas_encontradas.xlsx (gerado pelo script1)

Saídas:
  - arquivos/api_teste/guias_parcelamento/*.pdf
  - atualização do parcelas_encontradas.xlsx (EMITIDO/DATA_EMISSAO/CAMINHO_ARQUIVO)
  - atualização do controle_uso_api.xlsx

Regras:
  - não emite se EMITIDO=True
  - não emite se já houver SUCESSO de EMISSAO_GUIA_API no controle
  - imprime ERRO no console com DETALHE humano
  - decodifica PDF base64 de forma robusta
"""

import os
import sys
import json
import base64
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

ARQ_PARCELAS_ENCONTRADAS = PASTA_ARQUIVOS_API / "parcelas_encontradas.xlsx"
ARQ_CONTROLE = PASTA_ARQUIVOS_API / "controle_uso_api.xlsx"
PASTA_GUIAS = PASTA_ARQUIVOS_API / "guias_parcelamento"
PASTA_GUIAS.mkdir(parents=True, exist_ok=True)

# =========================
# CONFIG INTEGRA
# =========================

BASE_URL = "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1"

CNPJ_CONTRATANTE = os.getenv("CNPJ_CONTRATANTE", "").strip()
CNPJ_CONTRATANTE = re.sub(r"\D", "", CNPJ_CONTRATANTE).zfill(14)
if not CNPJ_CONTRATANTE or len(CNPJ_CONTRATANTE) != 14:
    raise RuntimeError("Defina CNPJ_CONTRATANTE no .env (14 dígitos).")

ID_SERVICOS_EMISSAO: Dict[str, str] = {
    "PARCSN": "GERARDAS161",
    "PARCSN-ESP": "GERARDAS171",
    "PERTSN": "GERARDAS181",
    "RELPSN": "GERARDAS191",
}

# =========================
# HELPERS
# =========================


def normalizar_cpf_cnpj(valor: Any) -> Tuple[str, int]:
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


def formatar_mensagens_humanas(mensagens: List[str]) -> str:
    limpas: List[str] = []
    for m in mensagens or []:
        if not m:
            continue
        s = str(m).replace("\n", " ").replace("\r", " ").strip()
        s = re.sub(r"\s+", " ", s)
        limpas.append(s)
    return " ; ".join(limpas)


def esta_emitido_flag(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    s = str(v).strip().upper()
    return s in {"TRUE", "1", "SIM", "S", "YES"}


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


def chaves_emissao_sucesso(dfc: pd.DataFrame) -> set:
    s = set()
    if dfc.empty:
        return s
    mask = (dfc["TIPO_OPERACAO"].astype(str).str.upper() == "EMISSAO_GUIA_API") & (
        dfc["RESULTADO"].astype(str).str.upper() == "SUCESSO"
    )
    for _, r in dfc[mask].iterrows():
        sistema = str(r.get("SISTEMA_SN", "")).strip().upper()
        doc = normalizar_cpf_cnpj(r.get("CPF_CNPJ", ""))[0]
        num = str(r.get("NUMERO_PARCELA", "")).strip()
        if not sistema or not doc or not num:
            continue
        try:
            num_i = int(re.sub(r"\D", "", num))
        except Exception:
            continue
        s.add((sistema, doc, num_i))
    return s


def buscar_pdf_b64(obj: Any) -> Optional[str]:
    if isinstance(obj, dict):
        for k, v in obj.items():
            lk = str(k).lower()
            if "pdf" in lk and "b64" in lk:
                return str(v)
        for v in obj.values():
            r = buscar_pdf_b64(v)
            if r:
                return r
    elif isinstance(obj, list):
        for it in obj:
            r = buscar_pdf_b64(it)
            if r:
                return r
    elif isinstance(obj, str):
        t = obj.strip()
        if t.startswith("JVBERi0"):  # assinatura típica base64 de PDF
            return t
    return None


def emitir_documento_arrecadacao(sistema_sn: str, cpf_cnpj: str, parcela_aaaamm: int) -> Dict[str, Any]:
    sistema_sn = sistema_sn.strip().upper()
    id_servico = ID_SERVICOS_EMISSAO.get(sistema_sn)
    if not id_servico:
        return {"ok": False, "mensagens": [f"idServico não configurado para {sistema_sn}"], "dados_raw": None, "pdf_b64": None}

    doc, tipo = normalizar_cpf_cnpj(cpf_cnpj)
    if not doc:
        return {"ok": False, "mensagens": ["CPF/CNPJ inválido após normalização."], "dados_raw": None, "pdf_b64": None}

    access_token, jwt_token = obter_tokens()

    payload = {
        "contratante": {"numero": CNPJ_CONTRATANTE, "tipo": 2},
        "autorPedidoDados": {"numero": CNPJ_CONTRATANTE, "tipo": 2},
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

    resp = requests.post(
        f"{BASE_URL}/Emitir",
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
        return {"ok": False, "mensagens": msgs, "dados_raw": raw, "pdf_b64": None}

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
        return {"ok": False, "mensagens": msgs + ["Resposta sem campo 'dados'."], "dados_raw": raw, "pdf_b64": None}

    if isinstance(dados_str, str):
        try:
            dados = json.loads(dados_str)
        except Exception:
            dados = dados_str
    else:
        dados = dados_str

    pdf_b64 = buscar_pdf_b64(dados)
    if not pdf_b64:
        return {"ok": False, "mensagens": msgs + ["PDF base64 não localizado na resposta."], "dados_raw": dados, "pdf_b64": None}

    return {"ok": True, "mensagens": msgs, "dados_raw": dados, "pdf_b64": pdf_b64}


# =========================
# MAIN
# =========================


def main() -> None:
    if not ARQ_PARCELAS_ENCONTRADAS.exists():
        raise FileNotFoundError(
            f'Arquivo "{ARQ_PARCELAS_ENCONTRADAS}" não encontrado. Gere antes pelo script1.'
        )

    df = pd.read_excel(ARQ_PARCELAS_ENCONTRADAS, dtype=str, na_filter=False, engine="openpyxl")

    obrig = ["EMPRESA", "CPF_CNPJ", "SISTEMA_SN", "NUMERO_PARCELA", "VALOR_PARCELA"]
    for c in obrig:
        if c not in df.columns:
            raise KeyError(f'Coluna obrigatória "{c}" não encontrada em {ARQ_PARCELAS_ENCONTRADAS.name}.')

    # garante colunas
    if "EMITIDO" not in df.columns:
        df["EMITIDO"] = "FALSE"
    if "DATA_EMISSAO" not in df.columns:
        df["DATA_EMISSAO"] = ""
    if "CAMINHO_ARQUIVO" not in df.columns:
        df["CAMINHO_ARQUIVO"] = ""

    # normaliza
    df["SISTEMA_SN"] = df["SISTEMA_SN"].astype(str).str.strip().str.upper()
    df["CPF_CNPJ"] = df["CPF_CNPJ"].apply(lambda v: normalizar_cpf_cnpj(v)[0])
    df["NUMERO_PARCELA"] = df["NUMERO_PARCELA"].astype(str).str.strip()

    # controle
    dfc = carregar_controle()
    ja_ok = chaves_emissao_sucesso(dfc)

    # pendentes
    pend_idx: List[int] = []
    for idx, r in df.iterrows():
        if esta_emitido_flag(r.get("EMITIDO", "")):
            continue
        sistema = str(r["SISTEMA_SN"]).strip().upper()
        doc = str(r["CPF_CNPJ"]).strip()
        num_s = str(r["NUMERO_PARCELA"]).strip()
        try:
            num_i = int(re.sub(r"\D", "", num_s))
        except Exception:
            continue
        if (sistema, doc, num_i) in ja_ok:
            continue
        pend_idx.append(idx)

    if not pend_idx:
        print("Nenhuma parcela pendente de emissão encontrada.")
        return

    print("Parcelas pendentes de emissão:")
    for idx in pend_idx:
        r = df.loc[idx]
        print(f"  - {r['SISTEMA_SN']} | {r['CPF_CNPJ']} | parcela {r['NUMERO_PARCELA']} | valor {r['VALOR_PARCELA']}")

    novos_logs: List[Dict[str, Any]] = []
    sucesso = 0
    erro = 0

    for idx in pend_idx:
        r = df.loc[idx]
        sistema = str(r["SISTEMA_SN"]).strip().upper()
        doc = str(r["CPF_CNPJ"]).strip()
        num_s = str(r["NUMERO_PARCELA"]).strip()

        try:
            num_i = int(re.sub(r"\D", "", num_s))
        except Exception:
            erro += 1
            msg = "NUMERO_PARCELA inválido."
            print(f"\nEmitindo guia para {sistema} | {doc} | parcela {num_s}...\n  -> ERRO: {msg}")
            novos_logs.append(
                {
                    "TIPO_OPERACAO": "EMISSAO_GUIA_API",
                    "SISTEMA_SN": sistema,
                    "CPF_CNPJ": doc,
                    "ID_PARCELAMENTO": "",
                    "ID_PARCELA_API": "",
                    "NUMERO_PARCELA": "",
                    "DATA_HORA": dt.datetime.now().isoformat(timespec="seconds"),
                    "RESULTADO": "ERRO",
                    "DETALHE": msg,
                    "DADOS_BRUTOS_JSON": "",
                }
            )
            continue

        print(f"\nEmitindo guia para {sistema} | {doc} | parcela {num_i}...")

        try:
            ret = emitir_documento_arrecadacao(sistema, doc, num_i)
        except Exception as e:
            ret = {"ok": False, "mensagens": [str(e)], "dados_raw": None, "pdf_b64": None}

        ok = bool(ret.get("ok"))
        msgs = ret.get("mensagens", [])
        dados_raw = ret.get("dados_raw")
        pdf_b64 = ret.get("pdf_b64")

        detalhe = formatar_mensagens_humanas(msgs)

        if ok and pdf_b64:
            try:
                b64_clean = "".join(str(pdf_b64).split())
                pdf_bytes = base64.b64decode(b64_clean, validate=False)
                if not pdf_bytes.startswith(b"%PDF"):
                    raise ValueError("Conteúdo decodificado não parece PDF (%PDF).")
            except Exception as e:
                ok = False
                erro += 1
                detalhe = formatar_mensagens_humanas(msgs + [f"Falha ao decodificar PDF base64: {e}"])
                print(f"  -> ERRO: {detalhe}")
                caminho_pdf = ""
            else:
                nome_pdf = f"{sistema}_{doc}_{num_i}.pdf".replace("/", "-").replace("\\", "-")
                caminho = PASTA_GUIAS / nome_pdf
                with open(caminho, "wb") as f:
                    f.write(pdf_bytes)

                df.at[idx, "EMITIDO"] = "TRUE"
                df.at[idx, "DATA_EMISSAO"] = dt.datetime.now().isoformat(timespec="seconds")
                df.at[idx, "CAMINHO_ARQUIVO"] = str(caminho)

                sucesso += 1
                caminho_pdf = str(caminho)
                print(f"  -> SUCESSO: {caminho_pdf}")
        else:
            erro += 1
            print(f"  -> ERRO: {detalhe}")
            caminho_pdf = ""

        novos_logs.append(
            {
                "TIPO_OPERACAO": "EMISSAO_GUIA_API",
                "SISTEMA_SN": sistema,
                "CPF_CNPJ": doc,
                "ID_PARCELAMENTO": "",
                "ID_PARCELA_API": "",
                "NUMERO_PARCELA": str(num_i),
                "DATA_HORA": dt.datetime.now().isoformat(timespec="seconds"),
                "RESULTADO": "SUCESSO" if (ok and pdf_b64) else "ERRO",
                "DETALHE": detalhe,
                "DADOS_BRUTOS_JSON": json.dumps(dados_raw, ensure_ascii=False) if dados_raw is not None else "",
            }
        )

    # grava controle
    if novos_logs:
        dfc = pd.concat([dfc, pd.DataFrame(novos_logs)], ignore_index=True)
        dfc.to_excel(ARQ_CONTROLE, index=False)

    # grava parcelas
    df.to_excel(ARQ_PARCELAS_ENCONTRADAS, index=False)

    print("\n=== RESUMO EMISSÃO (SCRIPT 2) ===")
    print(f"Total processadas: {len(pend_idx)}")
    print(f" - Sucesso: {sucesso}")
    print(f" - Erro:    {erro}")
    print(f"Controle: {ARQ_CONTROLE}")
    print(f"Planilha: {ARQ_PARCELAS_ENCONTRADAS}")
    print(f"PDFs: {PASTA_GUIAS}")


if __name__ == "__main__":
    main()
