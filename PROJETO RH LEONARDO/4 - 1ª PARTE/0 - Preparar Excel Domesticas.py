import argparse
import os
import re
import unicodedata
from pathlib import Path

try:
    import pandas as pd
except Exception as exc:
    raise SystemExit("pandas is required to read the Excel file") from exc


SOURCE_PATH = r"W:\DOCUMENTOS ESCRITORIO\RH\RH\POLLYANA\DOMESTICAS\RELAÇÃO DE EMPRESAS - DOMESTICAS.xlsx"
SHEET_NAME = "Domésticas"

OUTPUT_DIR = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\4 - 1ª PARTE"
OUTPUT_BASENAME = "domesticas_preparado.xlsx"


def is_number(value) -> bool:
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    text = str(value).strip()
    return text.isdigit()


def normalize_text(text: str) -> str:
    if text is None:
        raw = ""
    else:
        raw = str(text).strip()
    normalized = unicodedata.normalize("NFKD", raw)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch)).strip()


def normalize_header(text: str) -> str:
    clean = normalize_text(text).lower()
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean


def is_valid_name(name: str) -> bool:
    text = normalize_text(name).lower()
    if not text:
        return False
    return text.endswith(" - procuracao")


def read_source_with_detected_header(path: str, sheet: str) -> pd.DataFrame:
    preview = pd.read_excel(path, sheet_name=sheet, header=None)
    header_row = None
    for i in range(min(len(preview), 30)):
        row_values = [normalize_header(v) for v in preview.iloc[i].tolist()]
        joined = " | ".join(row_values)
        if "empregador" in joined and "vt/vr" in joined:
            header_row = i
            break
    if header_row is None:
        return pd.read_excel(path, sheet_name=sheet)
    return pd.read_excel(path, sheet_name=sheet, header=header_row)


def pick_column(df: pd.DataFrame, aliases: list[str], fallback_index: int):
    col_map = {normalize_header(c): c for c in df.columns}
    for alias in aliases:
        key = normalize_header(alias)
        if key in col_map:
            return col_map[key]
    if df.shape[1] <= fallback_index:
        return None
    return df.columns[fallback_index]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=SOURCE_PATH)
    ap.add_argument("--sheet", default=SHEET_NAME)
    ap.add_argument("--out-dir", default=OUTPUT_DIR)
    ap.add_argument("--out-name", default=OUTPUT_BASENAME)
    args = ap.parse_args()

    if not os.path.exists(args.source):
        raise SystemExit("Source Excel not found.")

    df = read_source_with_detected_header(args.source, args.sheet)

    if df.shape[1] < 6:
        raise SystemExit("Expected at least 6 columns (A-F).")

    a_col = pick_column(df, ["N°", "N", "No", "Numero"], 0)
    b_col = pick_column(df, ["Empregador", "Nome"], 1)
    c_col = pick_column(df, ["CPF"], 2)
    f_col = pick_column(df, ["VT/VR", "VT VR"], 5)
    if any(col is None for col in (a_col, b_col, c_col, f_col)):
        raise SystemExit("Nao foi possivel localizar as colunas N, Empregador/Nome, CPF e VT/VR.")

    df = df[df[a_col].apply(is_number)]
    df = df[df[b_col].apply(is_valid_name)]

    out_df = pd.DataFrame(
        {
            "N": df[a_col].astype(str).str.strip(),
            "Nome": df[b_col].astype(str).str.strip(),
            "CPF": df[c_col].astype(str).str.strip(),
            "VT/VR": df[f_col].astype(str).str.strip(),
        }
    )

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / args.out_name

    out_df.to_excel(out_path, index=False)
    print(f"OK: {out_path}")


if __name__ == "__main__":
    main()
