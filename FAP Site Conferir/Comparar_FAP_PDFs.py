import re
import sys
import zipfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

from pypdf import PdfReader

BASE_DIR = Path(__file__).resolve().parent
DOWNLOAD_DIR = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\FAP\2026")
BASE_PDF = BASE_DIR / "RelHistoricoRATFAP.pdf"
OUTPUT_XLSX = BASE_DIR / "inconsistencias_fap.xlsx"
OUTPUT_CSV = BASE_DIR / "inconsistencias_fap.csv"

CNPJ_RE = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
CNPJ_DIGITS_RE = re.compile(r"\d{14}")
NAME_CNPJ_RE = re.compile(
    r"(?:^|\n)\s*\d+\s*-\s*(?P<name>[^\n]+?)\s+(?P<cnpj>\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})",
    re.MULTILINE,
)
SEQ_RE = re.compile(
    r"(?P<cnpj>\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})"
    r"|(?P<vig>\d{2}/\d{4})\s*\d+%?\s*(?P<fap>\d{1,2}[.,]\d{4})\s*Vig.ncia:",
    re.MULTILINE,
)


@dataclass
class BaseEntry:
    name: str
    cnpj: str
    cnpj_digits: str
    vigencia: str
    fap: str


@dataclass
class SiteEntry:
    name: str
    cnpj: str
    cnpj_digits: str
    fap: str


def _read_pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def _vig_key(vigencia: str) -> tuple[int, int]:
    month, year = vigencia.split("/")
    return int(year), int(month)


def _to_decimal(value: str) -> Decimal | None:
    try:
        normalized = value.replace(".", "").replace(",", ".")
        return Decimal(normalized)
    except (InvalidOperation, AttributeError):
        return None


def _digits_only(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def _parse_base_pdf(path: Path) -> dict[str, BaseEntry]:
    text = _read_pdf_text(path)
    name_map: dict[str, str] = {}
    for m in NAME_CNPJ_RE.finditer(text):
        cnpj = m.group("cnpj")
        name = (m.group("name") or "").strip()
        if cnpj and name:
            name_map[cnpj] = name

    latest: dict[str, BaseEntry] = {}
    current_cnpj = ""
    for m in SEQ_RE.finditer(text):
        if m.group("cnpj"):
            current_cnpj = m.group("cnpj")
            continue
        if not current_cnpj:
            continue
        vig = m.group("vig") or ""
        fap = (m.group("fap") or "").strip()
        if not vig or not fap:
            continue
        cnpj_digits = _digits_only(current_cnpj)
        name = name_map.get(current_cnpj, "")
        entry = BaseEntry(
            name=name, cnpj=current_cnpj, cnpj_digits=cnpj_digits, vigencia=vig, fap=fap
        )
        if cnpj_digits not in latest:
            latest[cnpj_digits] = entry
            continue
        if _vig_key(vig) > _vig_key(latest[cnpj_digits].vigencia):
            latest[cnpj_digits] = entry
    return latest


def _parse_site_pdfs(folder: Path) -> dict[str, SiteEntry]:
    result: dict[str, SiteEntry] = {}
    for pdf_path in sorted(folder.glob("*.pdf")):
        try:
            text = _read_pdf_text(pdf_path)
        except Exception:
            continue
        cnpj_m = re.search(r"CNPJ\s*([0-9./-]+)", text)
        cnpj_digits = ""
        cnpj_formatted = ""
        if cnpj_m:
            cnpj_formatted = cnpj_m.group(1).strip()
            cnpj_digits = _digits_only(cnpj_formatted)
        if not cnpj_digits:
            digits_m = CNPJ_DIGITS_RE.search(text)
            if digits_m:
                cnpj_digits = digits_m.group(0)
                cnpj_formatted = cnpj_digits
        fap_m = re.search(r"Valor:\s*([0-9.,]+)", text)
        name_m = re.search(r"Raz[aã]o Social\s+([^\n\r]+)", text)
        if not cnpj_digits:
            continue
        name = (name_m.group(1).strip() if name_m else "")
        fap = (fap_m.group(1).strip() if fap_m else "")
        result[cnpj_digits] = SiteEntry(
            name=name, cnpj=cnpj_formatted, cnpj_digits=cnpj_digits, fap=fap
        )
    return result


def _write_csv(path: Path, headers: list[str], rows: Iterable[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        f.write(";".join(headers) + "\n")
        for row in rows:
            f.write(";".join(row) + "\n")


def _xml_escape(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _col_letter(idx: int) -> str:
    letters = ""
    n = idx + 1
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _build_sheet_xml(headers: list[str], rows: list[list[str]]) -> str:
    all_rows = [headers] + rows
    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "  <sheetData>",
    ]
    for r_idx, row in enumerate(all_rows, start=1):
        lines.append(f'    <row r="{r_idx}">')
        for c_idx, value in enumerate(row, start=1):
            cell = f"{_col_letter(c_idx - 1)}{r_idx}"
            safe = _xml_escape(value)
            lines.append(
                f'      <c r="{cell}" t="inlineStr"><is><t>{safe}</t></is></c>'
            )
        lines.append("    </row>")
    lines.append("  </sheetData>")
    lines.append("</worksheet>")
    return "\n".join(lines)


def _write_xlsx(path: Path, headers: list[str], rows: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    sheet_xml = _build_sheet_xml(headers, rows)
    workbook_xml = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Inconsistencias" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""
    rels_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"""
    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def main() -> None:
    if not BASE_PDF.exists():
        raise FileNotFoundError(f"PDF base nao encontrado: {BASE_PDF}")
    if not DOWNLOAD_DIR.exists():
        raise FileNotFoundError(f"Pasta de downloads nao encontrada: {DOWNLOAD_DIR}")

    base_map = _parse_base_pdf(BASE_PDF)
    site_map = _parse_site_pdfs(DOWNLOAD_DIR)

    rows: list[list[str]] = []
    for cnpj_digits, base in sorted(base_map.items()):
        name = base.name or site_map.get(cnpj_digits, SiteEntry("", "", cnpj_digits, "")).name
        fap_old = base.fap
        site = site_map.get(cnpj_digits)
        if not site or not site.fap:
            rows.append([name, base.cnpj, fap_old, "", "fora do site"])
            continue
        fap_new = site.fap
        if _to_decimal(fap_old) != _to_decimal(fap_new):
            rows.append([name, base.cnpj, fap_old, fap_new, "diferente"])
        else:
            rows.append([name, base.cnpj, fap_old, fap_new, "igual"])

    headers = ["nome", "cnpj", "fap_antigo", "fap_novo", "status"]
    _write_csv(OUTPUT_CSV, headers, rows)
    _write_xlsx(OUTPUT_XLSX, headers, rows)

    print(f"Total registros: {len(rows)}")
    print(f"CSV: {OUTPUT_CSV}")
    print(f"XLSX: {OUTPUT_XLSX}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"[erro] {exc}")
        sys.exit(1)
