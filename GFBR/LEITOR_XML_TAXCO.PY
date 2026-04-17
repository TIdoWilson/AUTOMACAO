#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lê XMLs de NFS-e e envia os dados extraídos diretamente para o Google Planilhas.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable, Optional
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog
import requests

# =========================
# Configurações do Google
# =========================
URL_WEB_APP = 'https://script.google.com/macros/s/AKfycbwLKUkUc3A_gxkPlC7E4MbnEKTSEJKiE6-1a5FdXOYKKd147VDMMqBqTo3SxyGNeAsD/exec'
TOKEN_SECRETO = '586431tokengfbrAFS3A4WASTSADGSGSDRD564'

# =========================
# Utilitários básicos
# =========================

def normalize_tag(tag: str) -> str:
    tag = tag.rsplit("}", 1)[-1]
    return re.sub(r"[^a-z0-9]", "", tag.lower())

@dataclass
class IndexedNode:
    tag: str
    text: str
    ancestors: tuple[str, ...]

@dataclass
class NotaRow:
    chave_unica: str
    mes_referencia: str
    numero_nfse: str
    numero_rps: str
    data_emissao: str
    cancelada: str
    cpf_cnpj_tomador: str
    nome_tomador: str
    valor_servicos: str
    valor_iss: str
    valor_ir: str
    valor_pis: str
    valor_cofins: str
    valor_csll: str
    valor_inss: str
    valor_liquido: str
    arquivo_xml: str
    data_importacao: str

def index_xml(root: ET.Element) -> list[IndexedNode]:
    indexed: list[IndexedNode] = []
    def walk(elem: ET.Element, ancestors: tuple[str, ...]) -> None:
        tag = normalize_tag(elem.tag)
        text = (elem.text or "").strip()
        if text:
            indexed.append(IndexedNode(tag=tag, text=text, ancestors=ancestors))
        new_ancestors = ancestors + (tag,)
        for child in elem:
            walk(child, new_ancestors)
    walk(root, tuple())
    return indexed

def find_first_text(nodes: Iterable[IndexedNode], tags: Iterable[str], ancestor_any: Optional[Iterable[str]] = None, ancestor_not: Optional[Iterable[str]] = None) -> str:
    tags_n = {normalize_tag(t) for t in tags}
    anc_any = {normalize_tag(a) for a in ancestor_any or []}
    anc_not = {normalize_tag(a) for a in ancestor_not or []}
    for node in nodes:
        if node.tag not in tags_n: continue
        anc_set = set(node.ancestors)
        if anc_any and not anc_set.intersection(anc_any): continue
        if anc_not and anc_set.intersection(anc_not): continue
        if node.text: return node.text
    return ""

def parse_date(value: str) -> Optional[datetime]:
    if not value: return None
    v = value.strip()
    if v.endswith("Z"): v = v[:-1] + "+00:00"
    formats = ["%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S.%f%z", "%d/%m/%Y", "%d/%m/%Y %H:%M:%S"]
    for fmt in formats:
        try: return datetime.strptime(v, fmt)
        except ValueError: continue
    m = re.match(r"^(\d{4}-\d{2}-\d{2})", v)
    if m:
        try: return datetime.strptime(m.group(1), "%Y-%m-%d")
        except ValueError: return None
    return None

def normalize_date_str(value: str) -> str:
    dt = parse_date(value)
    return dt.strftime("%d/%m/%Y") if dt else ""

def month_ref_from_date(value: str) -> str:
    dt = parse_date(value)
    return dt.strftime("%Y-%m") if dt else "SEM_DATA"

def clean_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")

def round_money(value: float) -> float:
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def format_currency(value: Optional[float]) -> str:
    if value is None: return ""
    return f"{value:.2f}".replace(".", ",")

def format_cnpj_text(value: str) -> str:
    digits = clean_digits(value)
    return f"'{digits}" if digits else ""

def parse_decimal(value: str) -> Optional[float]:
    if value is None: return None
    text = str(value).strip()
    if not text: return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."): text = text.replace(".", "").replace(",", ".")
        else: text = text.replace(",", "")
    elif "," in text: text = text.replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", ".", "-", "-."}: return None
    try: return float(text)
    except ValueError: return None

def configure_stdout_utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except AttributeError: pass

def selecionar_pasta_xml(xml_dir_inicial: Optional[Path] = None) -> Optional[Path]:
    raiz = tk.Tk()
    raiz.withdraw()
    raiz.attributes("-topmost", True)
    pasta_inicial = xml_dir_inicial if xml_dir_inicial and xml_dir_inicial.exists() else Path.cwd()
    pasta_selecionada = filedialog.askdirectory(title="Selecione a pasta onde estão os XML", initialdir=str(pasta_inicial), mustexist=True)
    raiz.destroy()
    if not pasta_selecionada: return None
    return Path(pasta_selecionada)


# =========================
# Extração de campos
# =========================

def extract_numero_nfse(nodes: list[IndexedNode]) -> str:
    attempts = [
        find_first_text(nodes, ["nNFSe", "numeroNFSe", "numero"], ancestor_any=["nfse", "infnfse"], ancestor_not=["rps", "identificacaorps"]),
        find_first_text(nodes, ["numero", "nNFSe"], ancestor_any=["compnfse", "nfse"], ancestor_not=["rps", "identificacaorps"]),
        find_first_text(nodes, ["nNFSe", "numeroNFSe", "numero"], ancestor_any=["dps", "infdps"], ancestor_not=["rps", "identificacaorps"]),
        find_first_text(nodes, ["numeroNFSe", "nNFSe"]),
    ]
    for val in attempts:
        if val: return val
    return ""

def extract_numero_rps(nodes: list[IndexedNode]) -> str:
    attempts = [
        find_first_text(nodes, ["numeroRps", "numero"], ancestor_any=["rps", "identificacaorps"]),
        find_first_text(nodes, ["numero"], ancestor_any=["rps"]),
        find_first_text(nodes, ["nDPS", "numeroDPS"], ancestor_any=["dps", "infdps"]),
    ]
    for val in attempts:
        if val: return val
    return ""

def is_cancelada(filename: str) -> str:
    name_lower = filename.lower()
    if "cancelad" in name_lower or "-canc" in name_lower or "_canc" in name_lower:
        return "sim"
    return "não"

def extract_data_emissao(nodes: list[IndexedNode]) -> str:
    attempts = [
        find_first_text(nodes, ["dataEmissao", "dhEmi", "dataHoraEmissao", "dhemi"], ancestor_any=["nfse", "infnfse"]),
        find_first_text(nodes, ["dhEmi", "dataEmissao", "dataHoraEmissao", "dhemi"], ancestor_any=["dps", "infdps"]),
        find_first_text(nodes, ["dataEmissao", "dhEmi", "dataHoraEmissao", "dhemi"]),
        find_first_text(nodes, ["competencia"]),
    ]
    for val in attempts:
        if val: return normalize_date_str(val)
    return ""

def extract_tomador_doc(nodes: list[IndexedNode]) -> str:
    cnpj = find_first_text(nodes, ["cnpj", "cnpjTomador", "cnpjcpf"], ancestor_any=["tomador", "tomadorservico", "destinatario", "tomadorservicos", "toma"])
    cpf = find_first_text(nodes, ["cpf", "cpfTomador", "cnpjcpf"], ancestor_any=["tomador", "tomadorservico", "destinatario", "tomadorservicos", "toma"])
    return clean_digits(cnpj or cpf)

def extract_tomador_nome(nodes: list[IndexedNode]) -> str:
    return find_first_text(nodes, ["razaoSocial", "xNome", "nome", "nomeRazaoSocial"], ancestor_any=["tomador", "tomadorservico", "destinatario", "tomadorservicos", "toma"]) or ""

def extract_value(nodes: list[IndexedNode], tags: list[str], preferred_ancestors: list[str] | None = None) -> Optional[float]:
    text = ""
    if preferred_ancestors: text = find_first_text(nodes, tags, ancestor_any=preferred_ancestors)
    if not text: text = find_first_text(nodes, tags)
    return parse_decimal(text)

def extract_valores(nodes: list[IndexedNode]) -> dict[str, Optional[float]]:
    valores = {
        "valor_servicos": extract_value(nodes, ["valorServicos", "valorServico", "vServPrest", "vServ", "valorTotalServicos"], ["valores", "servico", "valoresnfse", "servicos", "infdps"]),
        "valor_iss": extract_value(nodes, ["valorIss", "vIss", "iss", "valorISSRetido", "vISSQN", "vRetISSQN"], ["valores", "servico", "valoresnfse", "servicos", "infdps"]),
        "valor_ir": extract_value(nodes, ["valorIr", "vIr", "ir", "valorIRFonte", "vRetIRRF"], ["valores", "servico", "valoresnfse", "servicos", "tribfed"]),
        "valor_pis": extract_value(nodes, ["valorPis", "vPis", "pis", "vRetPis", "vRetPIS"], ["valores", "servico", "valoresnfse", "servicos", "tribfed"]),
        "valor_cofins": extract_value(nodes, ["valorCofins", "vCofins", "cofins", "vRetCofins", "vRetCOFINS"], ["valores", "servico", "valoresnfse", "servicos", "tribfed"]),
        "valor_csll": extract_value(nodes, ["valorCsll", "vCsll", "csll", "vRetCsll", "vRetCSLL", "valorCSLLRetido"], ["valores", "servico", "valoresnfse", "servicos", "tribfed"]),
        "valor_inss": extract_value(nodes, ["valorInss", "vInss", "inss", "vRetInss", "vRetINSS"], ["valores", "servico", "valoresnfse", "servicos", "tribfed"]),
        "valor_liquido": extract_value(nodes, ["valorLiquidoNfse", "valorLiquido", "vLiq", "valorLiquidoNFSe", "vTotNF"], ["valores", "servico", "valoresnfse", "servicos", "infdps"]),
    }
    base = valores["valor_servicos"] or 0.0
    csll_retido_xml = valores["valor_csll"] or 0.0
    if csll_retido_xml > 0:
        valores["valor_pis"] = round_money(base * 0.0065)
        valores["valor_cofins"] = round_money(base * 0.03)
        valores["valor_csll"] = round_money(base * 0.01)
    else:
        valores["valor_pis"] = 0.0
        valores["valor_cofins"] = 0.0
        valores["valor_csll"] = 0.0
    if valores["valor_liquido"] is None and valores["valor_servicos"] is not None:
        retencoes = sum(v or 0.0 for k, v in valores.items() if k.startswith("valor_") and k not in {"valor_servicos", "valor_liquido"})
        valores["valor_liquido"] = max(round_money((valores["valor_servicos"] or 0.0) - retencoes), 0.0)
    return valores

def build_unique_key(root: ET.Element, nodes: list[IndexedNode], xml_bytes: bytes, arquivo_nome: str) -> str:
    chave = find_first_text(nodes, ["chaveAcesso", "chave", "chaveNFSe"], ancestor_any=["nfse", "infnfse"])
    
    # Se a chave for muito curta, pode ser um ID genérico. Ignora para evitar sobrescrever notas.
    if chave and len(chave) < 15:
        chave = ""
        
    if not chave:
        id_attr = root.attrib.get("Id", "") or root.attrib.get("id", "")
        if len(id_attr) >= 15:
            chave = id_attr

    if chave:
        return chave.strip()

    # Fallback: Se não achou uma chave forte, obriga a usar o nome do arquivo na chave para garantir unicidade
    numero_nfse = extract_numero_nfse(nodes)
    return hashlib.sha1(f"{numero_nfse}_{arquivo_nome}".encode("utf-8")).hexdigest()

def parse_xml_file(xml_path: Path) -> NotaRow:
    xml_bytes = xml_path.read_bytes()
    root = ET.fromstring(xml_bytes)
    nodes = index_xml(root)

    data_emissao = extract_data_emissao(nodes)
    mes_referencia = month_ref_from_date(data_emissao)
    valores = extract_valores(nodes)
    data_importacao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cancelada = is_cancelada(xml_path.name)

    if cancelada == "sim":
        val_servicos = val_iss = val_ir = val_pis = val_cofins = val_csll = val_inss = val_liquido = "0,00"
    else:
        val_servicos = format_currency(valores["valor_servicos"])
        val_iss      = format_currency(valores["valor_iss"])
        val_ir       = format_currency(valores["valor_ir"])
        val_pis      = format_currency(valores["valor_pis"])
        val_cofins   = format_currency(valores["valor_cofins"])
        val_csll     = format_currency(valores["valor_csll"])
        val_inss     = format_currency(valores["valor_inss"])
        val_liquido  = format_currency(valores["valor_liquido"])

    return NotaRow(
        chave_unica=build_unique_key(root, nodes, xml_bytes, xml_path.name),
        mes_referencia=mes_referencia,
        numero_nfse=extract_numero_nfse(nodes),
        numero_rps=extract_numero_rps(nodes),
        data_emissao=data_emissao,
        cancelada=cancelada,
        cpf_cnpj_tomador=extract_tomador_doc(nodes),
        nome_tomador=extract_tomador_nome(nodes),
        valor_servicos=val_servicos,
        valor_iss=val_iss,
        valor_ir=val_ir,
        valor_pis=val_pis,
        valor_cofins=val_cofins,
        valor_csll=val_csll,
        valor_inss=val_inss,
        valor_liquido=val_liquido,
        arquivo_xml=xml_path.name,
        data_importacao=data_importacao,
    )

# =========================
# Memória e Integração
# =========================

CSV_HEADERS = [
    "chave_unica", "mes_referencia", "numero_nfse", "numero_rps", "data_emissao",
    "cancelada", "cpf_cnpj_tomador", "nome_tomador", "valor_servicos",
    "valor_iss", "valor_ir", "valor_pis", "valor_cofins", "valor_csll", "valor_inss",
    "valor_liquido", "arquivo_xml", "data_importacao",
]

def row_to_dict(row: NotaRow) -> dict[str, str]:
    data = {header: getattr(row, header) for header in CSV_HEADERS}
    data["cpf_cnpj_tomador"] = format_cnpj_text(data["cpf_cnpj_tomador"])
    return data

def collect_xml_files(xml_dir: Path) -> list[Path]:
    # Alterado para pegar .xml minúsculo e maiúsculo
    return sorted(p for p in xml_dir.rglob("*") if p.suffix.lower() == ".xml" and p.is_file())

def enviar_dados_planilha(rows: list[NotaRow], delimiter: str = ";") -> bool:
    if not rows: return False

    dedup: dict[str, NotaRow] = {}
    for row in rows:
        dedup[row.chave_unica] = row

    ordered_rows = sorted(
            dedup.values(),
            key=lambda r: (
                # Inverte a string de "dd/mm/aaaa" para "aaaammdd" internamente
                # para garantir que a ordem cronológica funcione perfeitamente
                f"{r.data_emissao[6:10]}{r.data_emissao[3:5]}{r.data_emissao[0:2]}" if r.data_emissao else "99999999",
                r.numero_nfse or "", 
                r.numero_rps or ""
            ),
        )

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=CSV_HEADERS, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)

    for row in ordered_rows:
        writer.writerow(row_to_dict(row))

    csv_string = output.getvalue()
    url_completa = f"{URL_WEB_APP}?token={TOKEN_SECRETO}"

    print(f"Enviando {len(ordered_rows)} notas para o Google Planilhas...")
    try:
        resposta = requests.post(url_completa, data=csv_string.encode("utf-8"), headers={"Content-Type": "text/plain"})
        print(f"Status: {resposta.status_code}")
        print(f"Resposta: {resposta.text}")
        return resposta.status_code == 200
    except Exception as e:
        print(f"Erro ao conectar com o Apps Script: {e}")
        return False

def processar_e_enviar(xml_dir: Path, delimiter: str = ";") -> list[tuple[Path, str]]:
    arquivos = collect_xml_files(xml_dir)
    if not arquivos:
        print(f"Nenhum XML encontrado em: {xml_dir}")
        return []

    notas_processadas = []
    erros: list[tuple[Path, str]] = []

    for xml_path in arquivos:
        try:
            row = parse_xml_file(xml_path)
            notas_processadas.append(row)
        except Exception as exc:
            erros.append((xml_path, str(exc)))

    if notas_processadas:
        enviar_dados_planilha(notas_processadas, delimiter)
    else:
        print("Nenhuma nota válida foi extraída.")

    return erros

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Lê XMLs e envia direto para o Google Planilhas.")
    parser.add_argument("--xml-dir", type=Path, help="Pasta raiz onde estão os XMLs.")
    return parser

def main() -> int:
    configure_stdout_utf8()
    args = build_parser().parse_args()

    xml_dir = args.xml_dir or selecionar_pasta_xml()
    if xml_dir is None:
        print("Seleção de pasta cancelada pelo usuário.")
        return 1

    if not xml_dir.exists() or not xml_dir.is_dir():
        print(f"Diretório de XML inválido: {xml_dir}")
        return 2

    erros = processar_e_enviar(xml_dir)

    if erros:
        print(f"\n[{len(erros)}] Arquivos com erro durante a extração (verifique se são NFS-e válidas):")
        # Mostrar só os primeiros 10 erros para não inundar a tela
        for path, err in erros[:10]:
            print(f"  - {path.name}: {err}")
        if len(erros) > 10:
            print(f"  ... e mais {len(erros) - 10} arquivos.")
        return 1

    return 0

if __name__ == "__main__":
    raise SystemExit(main())