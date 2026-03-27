#!/usr/bin/env python3

import argparse
import json
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

EXCLUDED_NAMES = {"manifest.json", "magalu_downloads.zip"}
ID_PATTERN = re.compile(r"_([0-9a-f]{24})\.([^.]+)$", re.IGNORECASE)
PDF_ASCII_DATE_PATTERN = re.compile(r"/CreationDate\s*\(D:(\d{14})")
PDF_HEX_DATE_PATTERN = re.compile(r"/CreationDate\s*<((?:[0-9A-Fa-f]{2})+)>")
DATE_RANGE_PATTERN = re.compile(r"(20\d{2})-\d{2}-\d{2}_a_(20\d{2})-\d{2}-\d{2}")
SHEET_NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
EXCEL_EPOCH = datetime(1899, 12, 30)


@dataclass
class FileDecision:
    path: Path
    kind: str
    year: str
    target_bucket: str
    duplicate_key: str | None
    metadata_id: str | None
    reason: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Analisa os downloads do Magalu, detecta duplicados e separa "
            "arquivos para manter, revisar ou descartar."
        )
    )
    parser.add_argument(
        "--source-dir",
        default=r"w:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\ALEATORIOS\magalu_downloads",
        help="Pasta com os arquivos baixados.",
    )
    parser.add_argument(
        "--catalog-json",
        default=r"C:\Users\Matheus\Downloads\magalu.json",
        help="JSON capturado da API /reports.",
    )
    parser.add_argument(
        "--output-dir",
        default="magalu_downloads_triagem",
        help="Pasta de saida para relatorios e, se usar --apply, para a triagem.",
    )
    parser.add_argument(
        "--years",
        nargs="+",
        default=["2025", "2026"],
        help="Anos que devem ser mantidos.",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Aplica a triagem copiando os arquivos para subpastas.",
    )
    parser.add_argument(
        "--move",
        action="store_true",
        help="Move os arquivos em vez de copiar. So faz sentido com --apply.",
    )
    return parser.parse_args()


def load_catalog(path: Path) -> dict[str, dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    return {
        item["id"]: item
        for item in payload
        if isinstance(item, dict) and isinstance(item.get("id"), str)
    }


def extract_metadata_id(path: Path) -> str | None:
    match = ID_PATTERN.search(path.name)
    return match.group(1) if match else None


def normalize_group_name(path: Path) -> str:
    return ID_PATTERN.sub(r".\2", path.name)


def extract_pdf_creation_date(path: Path) -> str | None:
    text = path.read_bytes().decode("latin-1", errors="ignore")

    match = PDF_ASCII_DATE_PATTERN.search(text)
    if match:
        return match.group(1)

    match = PDF_HEX_DATE_PATTERN.search(text)
    if not match:
        return None

    try:
        decoded = bytes.fromhex(match.group(1)).decode("latin-1", errors="ignore")
    except ValueError:
        return None

    decoded_match = re.search(r"D:(\d{14})", decoded)
    if decoded_match:
        return decoded_match.group(1)
    return None


def parse_sheet_rows(path: Path) -> list[list[str]]:
    with zipfile.ZipFile(path) as archive:
        xml_data = archive.read("xl/worksheets/sheet1.xml")

    tree = ET.fromstring(xml_data)
    rows: list[list[str]] = []

    for row in tree.findall(".//x:sheetData/x:row", SHEET_NS):
        values: list[str] = []
        for cell in row.findall("x:c", SHEET_NS):
            cell_type = cell.get("t")
            if cell_type == "inlineStr":
                values.append("".join(cell.itertext()).strip())
                continue

            value_node = cell.find("x:v", SHEET_NS)
            values.append((value_node.text if value_node is not None else "").strip())
        rows.append(values)

    return rows


def extract_sheet_year(path: Path, report: dict[str, Any] | None) -> str:
    range_match = DATE_RANGE_PATTERN.search(path.stem)
    if range_match:
        return range_match.group(1)

    parameters = report.get("parameters") if isinstance(report, dict) else None
    if isinstance(parameters, dict):
        for key in ("start_date", "end_date"):
            value = parameters.get(key)
            if isinstance(value, str) and re.match(r"20\d{2}-\d{2}-\d{2}", value):
                return value[:4]

    rows = parse_sheet_rows(path)
    if len(rows) < 2 or not rows[1]:
        return "unknown"

    first_value = rows[1][0]
    try:
        date_value = EXCEL_EPOCH + timedelta(days=float(first_value))
    except ValueError:
        return "unknown"
    return str(date_value.year)


def sheet_signature(path: Path) -> str:
    if path.suffix.lower() == ".csv":
        content = path.read_text(encoding="utf-8", errors="replace").splitlines()
        return json.dumps(content, ensure_ascii=False)
    return json.dumps(parse_sheet_rows(path), ensure_ascii=False)


def decide_files(
    source_dir: Path,
    catalog: dict[str, dict[str, Any]],
    allowed_years: set[str],
) -> tuple[list[FileDecision], dict[str, list[FileDecision]]]:
    decisions: list[FileDecision] = []
    groups: dict[str, list[FileDecision]] = defaultdict(list)

    for path in sorted(source_dir.iterdir()):
        if not path.is_file() or path.name in EXCLUDED_NAMES:
            continue

        metadata_id = extract_metadata_id(path)
        report = catalog.get(metadata_id or "")
        ext = path.suffix.lower()

        if ext == ".pdf":
            creation_date = extract_pdf_creation_date(path)
            year = creation_date[:4] if creation_date else "unknown"
            duplicate_key = (
                f"pdf::{normalize_group_name(path)}::{creation_date}"
                if creation_date
                else None
            )

            if year == "unknown":
                bucket = "revisar_pdf_sem_data"
                reason = "PDF sem CreationDate legivel"
            elif year not in allowed_years:
                bucket = "fora_periodo"
                reason = f"PDF fora do periodo alvo ({year})"
            else:
                bucket = "candidato_manter"
                reason = f"PDF do ano {year}"

        elif ext in {".xlsx", ".csv"}:
            year = extract_sheet_year(path, report)
            duplicate_key = f"sheet::{sheet_signature(path)}"

            if year not in allowed_years:
                bucket = "fora_periodo"
                reason = f"Planilha fora do periodo alvo ({year})"
            else:
                bucket = "candidato_manter"
                reason = f"Planilha do ano {year}"
        else:
            continue

        decision = FileDecision(
            path=path,
            kind=ext.lstrip("."),
            year=year,
            target_bucket=bucket,
            duplicate_key=duplicate_key,
            metadata_id=metadata_id,
            reason=reason,
        )
        decisions.append(decision)

        if duplicate_key:
            groups[duplicate_key].append(decision)

    for key, group in groups.items():
        if len(group) < 2:
            continue

        keeper = group[0]
        keeper.reason += " | primeiro item do grupo duplicado"

        for duplicate in group[1:]:
            duplicate.target_bucket = "duplicados"
            duplicate.reason = f"Duplicado de {keeper.path.name} via {key}"

    return decisions, groups


def build_summary(decisions: list[FileDecision]) -> dict[str, Any]:
    by_bucket = Counter(item.target_bucket for item in decisions)
    by_kind = Counter(item.kind for item in decisions)
    by_year = Counter(item.year for item in decisions)

    duplicates = [
        {
            "file": item.path.name,
            "reason": item.reason,
        }
        for item in decisions
        if item.target_bucket == "duplicados"
    ]

    review = [
        item.path.name
        for item in decisions
        if item.target_bucket == "revisar_pdf_sem_data"
    ]

    return {
        "total_files": len(decisions),
        "by_bucket": dict(by_bucket),
        "by_kind": dict(by_kind),
        "by_year": dict(by_year),
        "duplicates": duplicates,
        "review_pdf_without_date": review,
    }


def write_report(output_dir: Path, decisions: list[FileDecision], summary: dict[str, Any]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    payload = {
        "generated_at": datetime.now().isoformat(),
        "summary": summary,
        "files": [
            {
                "file": item.path.name,
                "bucket": item.target_bucket,
                "kind": item.kind,
                "year": item.year,
                "metadata_id": item.metadata_id,
                "reason": item.reason,
            }
            for item in decisions
        ],
    }

    (output_dir / "triagem_relatorio.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    lines = [
        f"Total analisado: {summary['total_files']}",
        f"Buckets: {summary['by_bucket']}",
        f"Tipos: {summary['by_kind']}",
        f"Anos: {summary['by_year']}",
        "",
        "Duplicados detectados:",
    ]
    lines.extend(f"- {item['file']} | {item['reason']}" for item in summary["duplicates"])
    lines.append("")
    lines.append("PDFs para revisao manual:")
    lines.extend(f"- {name}" for name in summary["review_pdf_without_date"])

    (output_dir / "triagem_resumo.txt").write_text(
        "\n".join(lines) + "\n",
        encoding="utf-8",
    )


def materialize_triage(
    output_dir: Path,
    decisions: list[FileDecision],
    move_files: bool,
) -> None:
    for decision in decisions:
        destination_dir = output_dir / decision.target_bucket
        destination_dir.mkdir(parents=True, exist_ok=True)
        destination = destination_dir / decision.path.name

        if move_files:
            shutil.move(str(decision.path), str(destination))
        else:
            shutil.copy2(decision.path, destination)


def main() -> int:
    args = parse_args()

    source_dir = Path(args.source_dir)
    catalog = load_catalog(Path(args.catalog_json))
    output_dir = Path(args.output_dir)

    decisions, _groups = decide_files(
        source_dir=source_dir,
        catalog=catalog,
        allowed_years=set(args.years),
    )
    summary = build_summary(decisions)
    write_report(output_dir, decisions, summary)

    print(f"Total analisado: {summary['total_files']}")
    print(f"Buckets: {summary['by_bucket']}")
    print(f"Anos: {summary['by_year']}")
    print(f"Duplicados: {len(summary['duplicates'])}")
    print(f"PDFs para revisao: {len(summary['review_pdf_without_date'])}")
    print(f"Relatorios em: {output_dir}")

    if args.apply:
        materialize_triage(
            output_dir=output_dir,
            decisions=decisions,
            move_files=args.move,
        )
        print("Triagem aplicada.")
    else:
        print("Dry-run concluido. Nenhum arquivo foi copiado ou movido.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
