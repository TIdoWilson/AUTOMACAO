import argparse
import csv
import re
from pathlib import Path
import shutil


def extract_code_from_filename(name: str) -> str:
    m = re.search(r"\s-\s*(\d+)\s-\s*folha", name, re.IGNORECASE)
    if not m:
        return ""
    try:
        return str(int(m.group(1)))
    except Exception:
        return m.group(1)


def list_codes_by_group(
    base: Path, year: str | None, month: str | None, grupos: list[str]
) -> dict[str, list[str]]:
    target = base
    if year:
        target = target / year
    if month:
        target = target / month
    found: dict[str, list[str]] = {}
    for grupo in grupos:
        folder = target / str(grupo)
        if not folder.exists():
            continue
        for pdf in folder.rglob("*.pdf"):
            code = extract_code_from_filename(pdf.name)
            if code:
                found.setdefault(grupo, [])
                if code not in found[grupo]:
                    found[grupo].append(code)
    return found


def parse_grupos(value: str | None) -> list[str]:
    if not value:
        return ["6", "13", "14"]
    parts = re.split(r"[;, ]+", value.strip())
    grupos = []
    for part in parts:
        if not part:
            continue
        if part not in ("6", "13", "14"):
            raise SystemExit("Grupo invalido. Validos: 6, 13, 14")
        if part not in grupos:
            grupos.append(part)
    if not grupos:
        raise SystemExit("Nenhum grupo informado.")
    return grupos


def resolve_latest_year_month(base: Path) -> tuple[str | None, str | None]:
    if not base.exists():
        return None, None
    years = [p.name for p in base.iterdir() if p.is_dir() and p.name.isdigit()]
    if not years:
        return None, None
    years.sort(key=lambda x: int(x))
    year = years[-1]
    months_base = base / year
    months = [p.name for p in months_base.iterdir() if p.is_dir() and p.name.isdigit()]
    if not months:
        return year, None
    months.sort(key=lambda x: int(x))
    return year, months[-1]


def load_estabelecimentos_from_csv(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    data: dict[str, str] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        _ = next(reader, None)
        for row in reader:
            if not row:
                continue
            code = (row[0] or "").strip()
            if not code.isdigit():
                continue
            estab = ""
            if len(row) >= 2:
                estab = (row[1] or "").strip()
            data[str(int(code))] = estab
    return data


def load_estabelecimentos_from_excel(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    try:
        import win32com.client as win32
    except Exception as exc:
        raise SystemExit("pywin32 is required to read the Excel file via COM.") from exc

    data: dict[str, set[str]] = {}
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(str(path), None, True)
    ws = wb.Worksheets.Item(1)
    used = ws.UsedRange
    rows = used.Rows.Count
    cols = used.Columns.Count

    header_row = 1
    header_col = None
    estab_col = None
    cnpj_col = None
    for r in range(1, min(50, rows) + 1):
        for c in range(1, cols + 1):
            val = str(used.Cells.Item(r, c).Text or "").strip().lower()
            if val == "código":
                header_row = r
                header_col = c
            if val == "estabelecimento":
                header_row = r
                estab_col = c
            if val == "cnpj/cpf/cei":
                header_row = r
                cnpj_col = c
        if header_col or estab_col:
            break

    code_col = 1
    if header_col and header_col > 1:
        code_col = header_col - 1
    name_col = header_col or 2
    if not estab_col:
        estab_col = 3
    if not cnpj_col:
        cnpj_col = 5

    current_code = None
    def _extract_cnpj_digits(cell_text: str, cell_val2) -> str:
        if isinstance(cell_val2, (int, float)):
            return str(int(cell_val2))
        if cell_val2 is not None:
            digits = re.sub(r"\D", "", str(cell_val2))
            if digits:
                return digits
        return re.sub(r"\D", "", cell_text or "")

    for r in range(header_row + 1, rows + 1):
        code_val = str(used.Cells.Item(r, code_col).Text or "").strip()
        if code_val.isdigit():
            current_code = str(int(code_val))
        if not current_code:
            continue
        cell = used.Cells.Item(r, cnpj_col)
        digits = _extract_cnpj_digits(str(cell.Text or "").strip(), cell.Value2)
        if not digits:
            continue
        if len(digits) >= 14:
            digits = digits[-14:]
        elif 12 <= len(digits) < 14:
            digits = digits.zfill(14)
        else:
            # CPF/CEI or invalid length
            continue
        estab = digits[8:12]
        try:
            estab = str(int(estab))
        except Exception:
            pass
        data.setdefault(current_code, set()).add(estab)

    wb.Close(False)
    excel.Quit()

    out: dict[str, str] = {}
    for code, estabs in data.items():
        out[code] = ",".join(sorted(estabs, key=lambda x: int(x) if x.isdigit() else x))
    return out


def format_base_excel(path: Path) -> Path:
    if not path.exists():
        return path
    try:
        import win32com.client as win32
    except Exception as exc:
        raise SystemExit("pywin32 is required to format the Excel file via COM.") from exc

    formatted_path = path.with_name("lista bruta - formatado.xlsx")
    shutil.copy2(path, formatted_path)

    excel = win32.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(str(formatted_path), None, False)
    ws = wb.Worksheets.Item(1)
    used = ws.UsedRange
    rows = used.Rows.Count
    cols = used.Columns.Count

    # Remove formulas by converting to values.
    try:
        used.Value2 = used.Value2
    except Exception:
        pass

    header_row = 1
    cnpj_col = None
    for r in range(1, min(50, rows) + 1):
        for c in range(1, cols + 1):
            val = str(used.Cells.Item(r, c).Text or "").strip().lower()
            if val == "cnpj/cpf/cei":
                header_row = r
                cnpj_col = c
                break
        if cnpj_col:
            break

    def _fmt_digits(digits: str) -> str:
        if len(digits) == 11:
            return f"{digits[0:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:11]}"
        if len(digits) >= 14:
            digits = digits[-14:]
            return f"{digits[0:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:14]}"
        if 12 <= len(digits) < 14:
            digits = digits.zfill(14)
            return f"{digits[0:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:14]}"
        return ""

    if cnpj_col:
        for r in range(header_row + 1, rows + 1):
            cell = used.Cells.Item(r, cnpj_col)
            text = str(cell.Text or "").strip()
            val2 = cell.Value2
            digits = ""
            if isinstance(val2, (int, float)):
                digits = str(int(val2))
            elif val2 is not None:
                digits = re.sub(r"\D", "", str(val2))
            if not digits:
                digits = re.sub(r"\D", "", text)
            if not digits:
                continue
            formatted = _fmt_digits(digits)
            if not formatted:
                continue
            cell.NumberFormat = "@"
            cell.Value2 = formatted

    wb.Save()
    wb.Close(False)
    excel.Quit()
    return formatted_path


def load_locks(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    data: dict[str, str] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        _ = next(reader, None)
        for row in reader:
            if not row:
                continue
            code = (row[0] or "").strip()
            if not code.isdigit():
                continue
            estab = ""
            if len(row) >= 2:
                estab = (row[1] or "").strip()
            data[str(int(code))] = estab
    return data


def normalize_estab_sequence(estab: str) -> str:
    if not estab:
        return ""
    parts = [p.strip() for p in estab.split(",") if p.strip().isdigit()]
    if not parts:
        return ""
    nums = sorted({int(p) for p in parts})
    return ",".join(str(i) for i in range(1, len(nums) + 1))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--base",
        default=r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\1ª PARTE",
        help="Pasta base do AUTOMATIZADO/1ª PARTE.",
    )
    ap.add_argument("--year", default=None, help="Ano (YYYY).")
    ap.add_argument("--month", default=None, help="Mes (MM).")
    ap.add_argument("--grupos", default=None, help="Grupos para ler (ex: 13;14).")
    ap.add_argument(
        "--output",
        default=None,
        help="Arquivo CSV de saida (padrao: empresas.csv na pasta do script).",
    )
    ap.add_argument(
        "--estab-from",
        default=str(Path(__file__).resolve().parent / "empresas.csv"),
        help="CSV de origem para preencher estabelecimentos (coluna 2).",
    )
    ap.add_argument(
        "--locks",
        default=str(Path(__file__).resolve().parent / "locks.csv"),
        help="CSV com estabelecimentos travados (tem prioridade).",
    )
    ap.add_argument(
        "--excel-bruto",
        default=str(Path(__file__).resolve().parent / "lista bruta.xlsx"),
        help="Excel bruto na pasta do script (padrao: lista bruta.xlsx).",
    )
    ap.add_argument(
        "--prefer-excel",
        action="store_true",
        help="(obsoleto) Mantido por compatibilidade. O Excel bruto sempre sera usado quando existir.",
    )
    args = ap.parse_args()

    base = Path(args.base)
    grupos = parse_grupos(args.grupos)
    year = args.year
    month = args.month
    if not year or not month:
        latest_year, latest_month = resolve_latest_year_month(base)
        year = year or latest_year
        month = month or latest_month
    codes_by_group = list_codes_by_group(base, year, month, grupos)
    estab_map = load_estabelecimentos_from_csv(Path(args.estab_from))
    excel_path = Path(args.excel_bruto)
    if excel_path.exists():
        formatted_path = format_base_excel(excel_path)
        estab_excel = load_estabelecimentos_from_excel(formatted_path)
        estab_map.update(estab_excel)
    locks_map = load_locks(Path(args.locks))
    output = Path(args.output) if args.output else Path(__file__).resolve().parent / "empresas.csv"
    with output.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["codigo", "estab"])
        for grupo in grupos:
            writer.writerow([f"GRUPO {grupo}", ""])
            codes = codes_by_group.get(grupo, [])
            for code in sorted(codes, key=lambda x: int(x)):
                code_key = str(int(code))
                estab_val = locks_map.get(code_key) or estab_map.get(code_key, "")
                writer.writerow([code, normalize_estab_sequence(estab_val)])
            writer.writerow([])
    print(f"[ok] salvo: {output}")


if __name__ == "__main__":
    main()
