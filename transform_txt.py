#!/usr/bin/env python3
"""
transform_txt.py
----------------
Script genérico para transformar um .txt de um layout "origem" (prefeitura) para um layout "destino" (seu ERP)
a partir de um arquivo de configuração YAML.

Uso:
    python transform_txt.py --config config.yaml --input origem.txt --output destino.txt

Requisitos:
    - Python 3.9+
    - pip install pandas pyyaml python-dateutil

Recursos suportados:
    - Layout de ORIGEM: "delimited" (csv/;|tab) ou "fixed_width" (larguras fixas)
    - Layout de DESTINO: "delimited" ou "fixed_width"
    - Mapeamento de campos por nome/índice, renomeação e ordem definidas
    - Transformações simples (funções prontas: strip, upper, lower, zfill, replace, date_parse, date_format, number_parse, number_format, concat, join)
    - Valores padrão (default) quando o campo de origem não existir
    - Validações (obrigatório, regex, conjunto permitido, comprimento)
    - Conversão de encoding (default: utf-8)

Autor: (gerado por assistente)
"""
from __future__ import annotations

import argparse
import sys
import re
import io
import csv
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Callable
from dateutil import parser as dateparser
from datetime import datetime
import pandas as pd

def _safe_get(d: Dict[str, Any], key: str, default=None):
    return d[key] if key in d else default

def t_strip(x: Any, args: Dict[str, Any] | None = None):
    return None if x is None else str(x).strip()

def t_upper(x: Any, args: Dict[str, Any] | None = None):
    return None if x is None else str(x).upper()

def t_lower(x: Any, args: Dict[str, Any] | None = None):
    return None if x is None else str(x).lower()

def t_zfill(x: Any, args: Dict[str, Any] | None = None):
    width = int(args.get("width", 0)) if args else 0
    return None if x is None else str(x).zfill(width)

def t_replace(x: Any, args: Dict[str, Any] | None = None):
    if x is None: return None
    frm = args.get("from", "")
    to = args.get("to", "")
    return str(x).replace(frm, to)

def t_date_parse(x: Any, args: Dict[str, Any] | None = None):
    if x in (None, "", "None"): return None
    dayfirst = bool(args.get("dayfirst", True)) if args else True
    yearfirst = bool(args.get("yearfirst", False)) if args else False
    return dateparser.parse(str(x), dayfirst=dayfirst, yearfirst=yearfirst)

def t_date_format(x: Any, args: Dict[str, Any] | None = None):
    if x in (None, "", "None"): return None
    fmt = args.get("format", "%Y-%m-%d") if args else "%Y-%m-%d"
    if isinstance(x, (datetime,)):
        dt = x
    else:
        dt = t_date_parse(x, args)
    return dt.strftime(fmt) if dt else None

def t_number_parse(x: Any, args: Dict[str, Any] | None = None):
    if x in (None, "", "None"): return None
    s = str(x)
    dec = args.get("decimal", ",") if args else ","
    grp = args.get("group", ".") if args else "."
    if grp:
        s = s.replace(grp, "")
    if dec and dec != ".":
        s = s.replace(dec, ".")
    try:
        return float(s)
    except ValueError:
        return None

def t_number_format(x: Any, args: Dict[str, Any] | None = None):
    if x in (None, "", "None"): return None
    n_dec = int(args.get("decimals", 2)) if args else 2
    dec = args.get("decimal", ",") if args else ","
    grp = args.get("group", ".") if args else "."
    try:
        v = float(x)
    except Exception:
        v = t_number_parse(x, args)
    if v is None:
        return None
    s = f"{v:,.{n_dec}f}"
    s = s.replace(",", "X").replace(".", dec).replace("X", grp)
    return s

def t_concat(values: List[Any], args: Dict[str, Any] | None = None):
    sep = args.get("sep", "") if args else ""
    parts = ["" if v is None else str(v) for v in values]
    return sep.join(parts)

def t_join(values: List[Any], args: Dict[str, Any] | None = None):
    return t_concat(values, args)

TRANSFORMS: Dict[str, Callable] = {
    "strip": t_strip,
    "upper": t_upper,
    "lower": t_lower,
    "zfill": t_zfill,
    "replace": t_replace,
    "date_parse": t_date_parse,
    "date_format": t_date_format,
    "number_parse": t_number_parse,
    "number_format": t_number_format,
    "concat": t_concat,
    "join": t_join,
}

def read_source(input_path: str, cfg: Dict[str, Any]) -> pd.DataFrame:
    enc = _safe_get(cfg, "encoding") or "utf-8"
    src = cfg["source"]
    s_type = src.get("type")
    if s_type == "delimited":
        delimiter = src.get("delimiter", ";")
        has_header = bool(src.get("header", False))
        quoting = csv.QUOTE_MINIMAL
        df = pd.read_csv(input_path, delimiter=delimiter, encoding=enc, header=0 if has_header else None, dtype=str, quoting=quoting, keep_default_na=False, na_values=[""])
        if not has_header:
            names = [f"col_{i}" for i in range(df.shape[1])]
            if "source_fields" in src:
                sf = src["source_fields"]
                if all("name" in c for c in sf) and len(sf) == len(names):
                    names = [c["name"] for c in sf]
            df.columns = names
        return df
    elif s_type == "fixed_width":
        widths = [int(c["width"]) for c in src["source_fields"]]
        names = [c["name"] for c in src["source_fields"]]
        df = pd.read_fwf(input_path, widths=widths, names=names, encoding=enc, dtype=str)
        return df
    else:
        raise ValueError("source.type deve ser 'delimited' ou 'fixed_width'")

def apply_mapping(df: pd.DataFrame, cfg: Dict[str, Any]) -> pd.DataFrame:
    mapping = cfg["mapping"]
    out_rows = []
    row_filters = cfg.get("row_filters", [])
    def row_passes(row: pd.Series) -> bool:
        for rf in row_filters:
            field = rf.get("field")
            op = rf.get("op", "eq")
            value = rf.get("value")
            rv = row.get(field)
            if op == "eq" and not (rv == value): return False
            if op == "ne" and not (rv != value): return False
            if op == "startswith" and not (str(rv).startswith(str(value))): return False
            if op == "not_startswith" and str(rv).startswith(str(value)): return False
            if op == "contains" and str(value) not in str(rv): return False
            if op == "not_contains" and str(value) in str(rv): return False
        return True

    for _, row in df.iterrows():
        if not row_passes(row):
            continue
        out_row = {}
        for target in mapping:
            t_name = target["target"]
            if "from" in target:
                origin = target["from"]
                if isinstance(origin, list):
                    vals = [row.get(o) for o in origin]
                else:
                    vals = row.get(origin)
            else:
                vals = None
            v = vals
            if "transform" in target:
                tr_list = target["transform"]
                for tr in tr_list:
                    fname = tr["fn"]
                    fargs = tr.get("args", {})
                    fn = TRANSFORMS.get(fname)
                    if not fn:
                        raise ValueError(f"Transform '{fname}' não suportada.")
                    if isinstance(v, list):
                        v = fn(v, fargs)
                    else:
                        v = fn(v, fargs)
            if (v is None or v == "") and "default" in target:
                v = target["default"]
            out_row[t_name] = v
        out_rows.append(out_row)
    return pd.DataFrame(out_rows)

def run_validations(df: pd.DataFrame, cfg: Dict[str, Any]):
    validations = cfg.get("validations", [])
    errors = []
    for rule in validations:
        field = rule["field"]
        kind = rule.get("type", "required")
        if kind == "required":
            missing = df[field].isna() | (df[field].astype(str) == "")
            idx = list(df.index[missing])
            if idx:
                errors.append(f"Campo obrigatório '{field}' vazio em linhas: {idx}")
        elif kind == "regex":
            pat = re.compile(rule["pattern"])
            bad = ~df[field].astype(str).apply(lambda s: bool(pat.fullmatch(s)))
            idx = list(df.index[bad])
            if idx:
                errors.append(f"Regex falhou em '{field}' (/{rule['pattern']}/) linhas: {idx}")
        elif kind == "in":
            allowed = set(rule["allowed"])
            bad = ~df[field].isin(allowed)
            idx = list(df.index[bad])
            if idx:
                errors.append(f"Valores inválidos em '{field}'. Permitidos: {sorted(allowed)}. Linhas: {idx}")
        elif kind == "length":
            minlen = int(rule.get("min", 0))
            maxlen = int(rule.get("max", 10**9))
            bad = ~df[field].astype(str).apply(lambda s: minlen <= len(s) <= maxlen)
            idx = list(df.index[bad])
            if idx:
                errors.append(f"Comprimento fora de faixa em '{field}' (min={minlen}, max={maxlen}) linhas: {idx}")
    if errors:
        joined = "\n - ".join(errors)
        raise ValueError("Falhas de validação:\n - " + joined)

def write_target(df: pd.DataFrame, output_path: str, cfg: Dict[str, Any]):
    enc = _safe_get(cfg, "encoding") or "utf-8"
    dst = cfg["target"]
    t_type = dst.get("type", "delimited")
    order = dst.get("order")
    if order:
        missing = [c for c in order if c not in df.columns]
        if missing:
            raise ValueError(f"Colunas faltando no resultado: {missing}")
        df = df[order]

    fillna_val = dst.get("fillna")
    if fillna_val is not None:
        df = df.fillna(fillna_val)
    else:
        df = df.fillna("")

    if t_type == "delimited":
        delimiter = dst.get("delimiter", ";")
        include_header = bool(dst.get("header", True))
        df.to_csv(output_path, sep=delimiter, index=False, header=include_header, encoding=enc, quoting=csv.QUOTE_MINIMAL)
    elif t_type == "fixed_width":
        fields = dst["fields"]
        lines = []
        for _, row in df.iterrows():
            parts = []
            for f in fields:
                name = f["name"]
                width = int(f["width"])
                align = f.get("align", "left")
                pad = f.get("pad", " ")
                val = "" if pd.isna(row[name]) else str(row[name])
                if len(val) > width:
                    val = val[:width]
                if align == "right":
                    parts.append(val.rjust(width, pad[0]))
                else:
                    parts.append(val.ljust(width, pad[0]))
            lines.append("".join(parts))
        with open(output_path, "w", encoding=enc, newline="") as f:
            for line in lines:
                f.write(line + "\n")
    else:
        raise ValueError("target.type deve ser 'delimited' ou 'fixed_width'")

def main():
    import yaml
    ap = argparse.ArgumentParser(description="Transforma txt de um layout para outro usando config YAML.")
    ap.add_argument("--config", required=True, help="Caminho do config YAML.")
    ap.add_argument("--input", required=True, help="Caminho do arquivo de origem (.txt).")
    ap.add_argument("--output", required=True, help="Caminho do arquivo de saída.")
    args = ap.parse_args()

    with open(args.config, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    df_src = read_source(args.input, cfg)
    df_map = apply_mapping(df_src, cfg)
    run_validations(df_map, cfg)
    write_target(df_map, args.output, cfg)
    print(f"OK: gerado '{args.output}' com {len(df_map)} linhas.")

if __name__ == "__main__":
    main()
