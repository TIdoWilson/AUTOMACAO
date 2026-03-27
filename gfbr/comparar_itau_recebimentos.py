from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path

import pandas as pd


MONTHS_PT = {
    "janeiro": 1,
    "fevereiro": 2,
    "marco": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}


def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().strip()
    return " ".join(text.split())


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={c: normalize_text(c) for c in df.columns})


def extract_period_from_filename(path: Path) -> tuple[int, int]:
    normalized_name = normalize_text(path.stem).replace("-", ".").replace("_", ".")

    month = None
    for month_name, month_number in MONTHS_PT.items():
        if re.search(rf"\b{month_name}\b", normalized_name):
            month = month_number
            break

    if month is not None:
        date_match = re.search(r"\b\d{1,2}[./]\d{1,2}[./](20\d{2})\b", normalized_name)
        if date_match:
            year = int(date_match.group(1))
            return year, month

        years = re.findall(r"\b(20\d{2})\b", normalized_name)
        if years:
            return int(years[-1]), month

        raise ValueError(f"Nao foi possivel identificar o ano no nome do arquivo: {path.name}")

    num_match = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[./](20\d{2})(?!\d)", normalized_name)
    if num_match:
        month = int(num_match.group(1))
        year = int(num_match.group(2))
        return year, month

    raise ValueError(f"Nao foi possivel identificar o mes/ano no nome do arquivo: {path.name}")



def select_input_file(title: str, initial_dir: Path, filetypes: list[tuple[str, str]]) -> Path:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(
        title=title,
        initialdir=str(initial_dir),
        filetypes=filetypes,
    )
    root.destroy()

    if not selected:
        raise SystemExit("Selecao de arquivo cancelada pelo usuario.")

    return Path(selected)


def select_output_file(title: str, initial_dir: Path, default_filename: str) -> Path:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.asksaveasfilename(
        title=title,
        initialdir=str(initial_dir),
        initialfile=default_filename,
        defaultextension=".xlsx",
        filetypes=[("Planilha Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
    )
    root.destroy()

    if not selected:
        raise SystemExit("Selecao de arquivo de saida cancelada pelo usuario.")

    return Path(selected)


def find_sheet_name(excel_path: Path, target_name: str, engine: str | None = None) -> str:
    workbook = pd.ExcelFile(excel_path, engine=engine)
    normalized_target = normalize_text(target_name)

    for sheet in workbook.sheet_names:
        if normalize_text(sheet) == normalized_target:
            return sheet

    raise ValueError(
        f"Aba '{target_name}' nao encontrada em {excel_path.name}. "
        f"Abas disponiveis: {workbook.sheet_names}"
    )


def parse_excel_date_column(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    numeric_ratio = float(numeric.notna().mean())
    serial_ratio = float(numeric.between(20000, 80000).mean())

    if numeric_ratio >= 0.70 and serial_ratio >= 0.70:
        return pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")

    return pd.to_datetime(series, errors="coerce")


def prepare_bancos_itau(path: Path, year: int, month: int) -> pd.DataFrame:
    sheet_name = find_sheet_name(path, "itau")
    df = pd.read_excel(path, sheet_name=sheet_name)
    df = normalize_columns(df)

    required_columns = [
        "data de lancamento",
        "c/d (ml)",
    ]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatorias ausentes em {path.name}: {missing}")

    df["data_lancamento"] = pd.to_datetime(df["data de lancamento"], errors="coerce")
    df["valor"] = pd.to_numeric(df["c/d (ml)"], errors="coerce").round(2)
    df["detalhes"] = df.get("detalhes", "")
    df["origem"] = df.get("origem", "")
    df["no_transacao"] = df.get("no transacao", "")
    df["no_origem"] = df.get("no origem", "")
    df["conta_contrapartida"] = df.get("conta de contrapartida", "")

    filtered = df[
        (df["data_lancamento"].dt.year == year)
        & (df["data_lancamento"].dt.month == month)
        & (df["valor"] > 0)
    ].copy()

    filtered["data"] = filtered["data_lancamento"].dt.normalize()
    filtered["valor"] = filtered["valor"].round(2)

    return filtered[
        [
            "data",
            "valor",
            "no_transacao",
            "no_origem",
            "origem",
            "conta_contrapartida",
            "detalhes",
        ]
    ]


def prepare_recebimentos(path: Path, year: int, month: int) -> pd.DataFrame:
    sheet_name = find_sheet_name(path, "recebimentos", engine="pyxlsb")
    df = pd.read_excel(path, sheet_name=sheet_name, engine="pyxlsb")
    df = normalize_columns(df)

    required_columns = [
        "data recebimento",
        "valor recebido (r$)",
        "banco",
    ]
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatorias ausentes em {path.name}: {missing}")

    df["data"] = parse_excel_date_column(df["data recebimento"]).dt.normalize()
    df["valor"] = pd.to_numeric(df["valor recebido (r$)"], errors="coerce").round(2)
    df["banco"] = df["banco"].astype(str)

    filtered = df[
        (df["data"].dt.year == year)
        & (df["data"].dt.month == month)
        & (df["valor"].notna())
        & (df["banco"].str.contains("itau", case=False, na=False))
    ].copy()

    filtered["numero_nf"] = filtered.get("numero nf", "")
    filtered["clientes"] = filtered.get("clientes", "")
    filtered["valor_liquido"] = pd.to_numeric(filtered.get("valor liquido (r$)"), errors="coerce")

    return filtered[
        [
            "data",
            "valor",
            "banco",
            "numero_nf",
            "clientes",
            "valor_liquido",
        ]
    ]


def conciliate(
    bancos_df: pd.DataFrame, receb_df: pd.DataFrame, year: int, month: int
) -> tuple[pd.DataFrame, pd.DataFrame]:
    bancos = bancos_df.copy()
    receb = receb_df.copy()

    bancos = bancos.sort_values(["data", "valor", "no_transacao"], na_position="last").reset_index(drop=True)
    receb = receb.sort_values(["data", "valor", "numero_nf"], na_position="last").reset_index(drop=True)

    bancos["_chave"] = bancos["data"].dt.strftime("%Y-%m-%d") + "|" + bancos["valor"].map(lambda x: f"{x:.2f}")
    receb["_chave"] = receb["data"].dt.strftime("%Y-%m-%d") + "|" + receb["valor"].map(lambda x: f"{x:.2f}")

    bancos["_ord"] = bancos.groupby("_chave").cumcount()
    receb["_ord"] = receb.groupby("_chave").cumcount()

    bancos_pref = bancos.add_prefix("bancos_")
    receb_pref = receb.add_prefix("receb_")

    merged = bancos_pref.merge(
        receb_pref,
        left_on=["bancos__chave", "bancos__ord"],
        right_on=["receb__chave", "receb__ord"],
        how="outer",
        indicator=True,
    )

    merged["status"] = merged["_merge"].map(
        {
            "both": "EM AMBAS",
            "left_only": "SOMENTE BANCOS SW",
            "right_only": "SOMENTE STURMER RECEBIMENTOS",
        }
    )

    merged["data_referencia"] = merged["bancos_data"].fillna(merged["receb_data"])
    merged["valor_referencia"] = merged["bancos_valor"].fillna(merged["receb_valor"])
    merged["periodo"] = f"{year:04d}-{month:02d}"

    ordered_columns = [
        "periodo",
        "status",
        "data_referencia",
        "valor_referencia",
        "bancos_no_transacao",
        "bancos_no_origem",
        "bancos_origem",
        "bancos_conta_contrapartida",
        "bancos_detalhes",
        "receb_banco",
        "receb_numero_nf",
        "receb_clientes",
        "receb_valor_liquido",
    ]

    comparison = merged[ordered_columns].sort_values(
        ["status", "data_referencia", "valor_referencia"], na_position="last"
    )

    summary = pd.DataFrame(
        [
            {
                "periodo": f"{year:04d}-{month:02d}",
                "status": status,
                "quantidade": int(group.shape[0]),
                "soma_valor": float(group["valor_referencia"].fillna(0).sum()),
            }
            for status, group in comparison.groupby("status", dropna=False, observed=False)
        ]
    ).sort_values("status")

    return comparison, summary


def save_output(
    output_path: Path,
    comparison: pd.DataFrame,
    summary: pd.DataFrame,
    bancos_file: Path,
    receb_file: Path,
) -> None:
    metadata = pd.DataFrame(
        [
            {"campo": "arquivo_bancos_sw", "valor": str(bancos_file)},
            {"campo": "arquivo_sturmer_recebimentos", "valor": str(receb_file)},
        ]
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        comparison.to_excel(writer, index=False, sheet_name="comparativo")
        comparison[comparison["status"] == "EM AMBAS"].to_excel(
            writer, index=False, sheet_name="em_ambas"
        )
        comparison[comparison["status"] == "SOMENTE BANCOS SW"].to_excel(
            writer, index=False, sheet_name="somente_bancos_sw"
        )
        comparison[comparison["status"] == "SOMENTE STURMER RECEBIMENTOS"].to_excel(
            writer, index=False, sheet_name="somente_sturmer"
        )
        summary.to_excel(writer, index=False, sheet_name="resumo")
        metadata.to_excel(writer, index=False, sheet_name="origem_arquivos")


def run(bancos_path: Path, recebimentos_path: Path, output_path: Path | None) -> Path:
    period_bancos = extract_period_from_filename(bancos_path)
    period_receb = extract_period_from_filename(recebimentos_path)

    if period_bancos != period_receb:
        raise ValueError(
            "Os arquivos tem periodos diferentes no nome. "
            f"Bancos SW: {period_bancos[1]:02d}/{period_bancos[0]} | "
            f"Sturmer: {period_receb[1]:02d}/{period_receb[0]}"
        )

    year, month = period_bancos

    bancos_df = prepare_bancos_itau(bancos_path, year, month)
    receb_df = prepare_recebimentos(recebimentos_path, year, month)

    comparison, summary = conciliate(bancos_df, receb_df, year, month)

    if output_path is None:
        output_path = bancos_path.parent / f"COMPARATIVO_ITAU_RECEBIMENTOS_{year:04d}-{month:02d}.xlsx"

    save_output(output_path, comparison, summary, bancos_path, recebimentos_path)
    return output_path


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Concilia entradas do ITAU (BANCOS SW) com a aba RECEBIMENTOS (STURMER)."
    )
    parser.add_argument(
        "--bancos",
        type=Path,
        default=None,
        help="Caminho da planilha BANCOS SW (.xlsx). Se nao informar, abre janela para selecionar.",
    )
    parser.add_argument(
        "--recebimentos",
        type=Path,
        default=None,
        help="Caminho da planilha STURMER RECEBIMENTOS (.xlsb). Se nao informar, abre janela para selecionar.",
    )
    parser.add_argument(
        "--saida",
        type=Path,
        default=None,
        help="Caminho do arquivo de saida (.xlsx). Se nao informar, abre janela para selecionar.",
    )
    return parser


def main() -> None:
    parser = build_argument_parser()
    args = parser.parse_args()

    default_dir = Path(__file__).resolve().parent

    bancos_path = args.bancos
    if bancos_path is None:
        bancos_path = select_input_file(
            title="Selecione a planilha BANCOS SW",
            initial_dir=default_dir,
            filetypes=[("Planilhas Excel", "*.xlsx *.xlsm *.xls"), ("Todos os arquivos", "*.*")],
        )

    recebimentos_path = args.recebimentos
    if recebimentos_path is None:
        recebimentos_path = select_input_file(
            title="Selecione a planilha STURMER RECEBIMENTOS",
            initial_dir=bancos_path.parent,
            filetypes=[("Planilhas Excel Binarias", "*.xlsb"), ("Todos os arquivos", "*.*")],
        )

    output_path = args.saida
    if output_path is None:
        try:
            year, month = extract_period_from_filename(bancos_path)
            default_name = f"COMPARATIVO_ITAU_RECEBIMENTOS_{year:04d}-{month:02d}.xlsx"
        except Exception:
            default_name = "COMPARATIVO_ITAU_RECEBIMENTOS.xlsx"

        output_path = select_output_file(
            title="Salvar comparativo como",
            initial_dir=bancos_path.parent,
            default_filename=default_name,
        )

    output_path = run(bancos_path, recebimentos_path, output_path)
    print(f"Comparativo gerado em: {output_path}")


if __name__ == "__main__":
    main()
