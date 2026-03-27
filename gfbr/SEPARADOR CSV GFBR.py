# -*- coding: utf-8 -*-
"""
Processa arquivos Excel da aba 'baixa CSV', usando a coluna 'Competência' para separar por ano
e gerar arquivos CSV de até 50 linhas, na mesma pasta do arquivo original.

- Colunas C e G: formato "dd/mm/aaaa" sem horas
- Colunas F, H, I, J, K, L, M, N: decimais com vírgula

Requisitos:
    pip install pandas openpyxl
"""

from __future__ import annotations
import math
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog
from typing import List, Optional, Tuple, Dict
import pandas as pd

# --------- CONFIG FIXA --------- #
PASTA_ORIGEM = Path(r"W:\PASTA CLIENTES\TAXCO EDUCAÇÃO LTDA\CONCILIACAO\2026")
SHEET_NAME = "BAIXAS"
YEAR_SOURCE_COLUMN = "DATA EMISSÃO"
MAX_LINHAS_POR_ARQUIVO = 50
CSV_SEP = ";"
VALID_EXTS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
# Mapeamento das colunas por LETRA -> índice (0-based)
COL_C = 2   # terceira coluna
COL_G = 6   # sétima coluna
DECIMAL_COLS = [5, 7, 8, 9, 10, 11, 12, 13]  # F,H,I,J,K,L,M,N (0-based)
# -------------------------------- #

def log(msg: str) -> None:
    print(msg, file=sys.stdout, flush=True)

def is_hidden_excel_temp(p: Path) -> bool:
    return p.name.startswith("~$")

def find_excel_files(root: Path) -> List[Path]:
    return sorted([p for p in root.glob("*") if p.is_file() and p.suffix.lower() in VALID_EXTS and not is_hidden_excel_temp(p)])

def normalize_colname(name: str) -> str:
    return str(name).strip().lower()

def _series_to_year(series: pd.Series) -> Optional[pd.Series]:
    s = pd.to_numeric(series, errors="coerce").astype("Int64")
    valid = (s >= 1900) & (s <= 2100)
    if valid.sum() == 0:
        return None
    return s.astype("Int64")

def _series_to_year_from_date(series: pd.Series) -> Optional[pd.Series]:
    dt = pd.to_datetime(series, errors="coerce", dayfirst=True)
    if dt.notna().sum() == 0:
        return None
    return dt.dt.year.astype("Int64")

def try_year_from_forced_col(df: pd.DataFrame, forced_col: str) -> Tuple[pd.Series, str]:
    cols = list(df.columns)
    norm_forced = normalize_colname(forced_col)
    col = forced_col if forced_col in df.columns else next((c for c in cols if normalize_colname(c) == norm_forced), None)
    if not col:
        raise ValueError(f"Coluna '{forced_col}' não encontrada.")
    yr = _series_to_year(df[col])
    if yr is None:
        yr = _series_to_year_from_date(df[col])
    if yr is None:
        raise ValueError(f"Não foi possível extrair ANO da coluna '{col}'.")
    return yr, col

def chunk_dataframe(df: pd.DataFrame, size: int):
    if size <= 0:
        return [df]
    n = int(math.ceil(len(df) / float(size)))
    return [df.iloc[i*size:(i+1)*size] for i in range(n)]

def safe_stem(p: Path) -> str:
    return p.stem.replace(" ", "_").replace("/", "_").replace("\\", "_")[:120]

def format_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Formatar colunas C e G (por índice) para dd/mm/aaaa
    for col_idx in [COL_C, COL_G]:
        if col_idx < len(df.columns):
            col_name = df.columns[col_idx]
            df[col_name] = pd.to_datetime(df[col_name], errors="coerce", dayfirst=True)
            df[col_name] = df[col_name].dt.strftime("%d/%m/%Y")

    # Formatar colunas decimais (F,H,I,J,K,L,M,N)
    for col_idx in DECIMAL_COLS:
        if col_idx < len(df.columns):
            col_name = df.columns[col_idx]
            df[col_name] = pd.to_numeric(df[col_name], errors="coerce")
            df[col_name] = df[col_name].map(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(x) else "")

    return df

def write_csv(df: pd.DataFrame, out_path: Path, sep: str = CSV_SEP) -> None:
    df.to_csv(out_path, sep=sep, index=False, encoding="utf-8-sig")

def select_excel_file() -> Optional[Path]:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel para processar",
        filetypes=[
            ("Arquivos Excel", "*.xlsx *.xlsm *.xlsb *.xls"),
            ("Todos os arquivos", "*.*"),
        ],
    )
    root.destroy()
    if not arquivo:
        return None
    return Path(arquivo)

def process_excel(path: Path) -> Dict[str, int]:
    log(f"\n>>> Lendo: {path.name}")
    engine = "openpyxl" if path.suffix.lower() in {".xlsx", ".xlsm"} else None
    try:
        df = pd.read_excel(path, sheet_name=SHEET_NAME, engine=engine, dtype=object)
    except Exception as e:
        log(f"   [ERRO] Falha ao ler '{path.name}' (aba: {SHEET_NAME}). Motivo: {e}")
        return {}
    df = df.dropna(how="all").copy()
    if df.empty:
        log("   [AVISO] Aba sem dados.")
        return {}
    try:
        anos_series, col_usada = try_year_from_forced_col(df, YEAR_SOURCE_COLUMN)
        df["__ANO__"] = anos_series
        df = df.dropna(subset=["__ANO__"])
        df["__ANO__"] = df["__ANO__"].astype("Int64")
        log(f"   Coluna usada para ano: {col_usada}")
    except ValueError as e:
        log(f"   [ERRO] {e}")
        return {}
    if df.empty:
        log("   [AVISO] Após extrair ano, não restaram linhas válidas.")
        return {}
    out_counts: Dict[str, int] = {}
    for ano, df_ano in df.groupby("__ANO__", dropna=True):
        ano_int = int(ano)
        df_ano = df_ano.drop(columns=["__ANO__"])
        df_ano = format_columns(df_ano)  # aplica formatações
        partes = chunk_dataframe(df_ano, MAX_LINHAS_POR_ARQUIVO)
        stem = safe_stem(path)
        for idx, pedaco in enumerate(partes, start=1):
            out_name = f"{stem}__{ano_int}__parte-{idx:02d}.csv"
            out_path = path.parent / out_name
            try:
                write_csv(pedaco, out_path, sep=CSV_SEP)
                log(f"   [OK] Gerado: {out_name} ({len(pedaco)} linhas)")
                out_counts[str(ano_int)] = out_counts.get(str(ano_int), 0) + len(pedaco)
            except Exception as e:
                log(f"   [ERRO] Falha ao escrever '{out_name}': {e}")
    return out_counts

def main():
    arquivo = select_excel_file()
    if not arquivo:
        log("[AVISO] Nenhum arquivo selecionado.")
        sys.exit(0)
    if arquivo.suffix.lower() not in VALID_EXTS:
        log(f"[ERRO] Arquivo inválido. Selecione um Excel: {arquivo.name}")
        sys.exit(1)
    total_resumo = process_excel(arquivo)
    if total_resumo:
        log("\n===== RESUMO GERAL =====")
        for ano in sorted(total_resumo.keys()):
            log(f"{ano}: {total_resumo[ano]} linhas")
    else:
        log("\nNenhum arquivo gerado.")

if __name__ == "__main__":
    main()

