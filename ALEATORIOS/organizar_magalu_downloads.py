#!/usr/bin/env python3

import argparse
import json
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

EXCLUDED_NAMES = {"manifest.json", "magalu_downloads.zip"}
ID_PATTERN = re.compile(r"_([0-9a-f]{24})\.([^.]+)$", re.IGNORECASE)
TOKEN_PATTERN = re.compile(r"^(.*)_(pdf|xlsx|csv)_[0-9a-f]{24}$", re.IGNORECASE)
PDF_ASCII_DATE_PATTERN = re.compile(r"/CreationDate\s*\(D:(\d{14})")
PDF_HEX_DATE_PATTERN = re.compile(r"/CreationDate\s*<((?:[0-9A-Fa-f]{2})+)>")
DATE_RANGE_PATTERN = re.compile(r"(20\d{2})-(\d{2})-\d{2}_a_(20\d{2})-(\d{2})-\d{2}")
SHEET_NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
EXCEL_EPOCH = datetime(1899, 12, 30)


@dataclass
class FileEntry:
    path: Path
    report_id: str | None
    report: dict[str, Any] | None
    ext: str
    month_label: str
    base_name: str
    duplicate_group: str | None
    sort_key: tuple[Any, ...]
    bucket: str = "manter"
    target_name: str = ""
    uses_month_suffix: bool = False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Copia os downloads do Magalu para uma pasta organizada, removendo "
            "duplicados por criterio seguro e renomeando os arquivos por mes."
        )
    )
    parser.add_argument(
        "--source-dir",
        default=r"w:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\ALEATORIOS\magalu_downloads",
        help="Pasta original com os downloads.",
    )
    parser.add_argument(
        "--catalog-json",
        default=r"C:\Users\Matheus\Downloads\magalu.json",
        help="JSON capturado da API /reports.",
    )
    parser.add_argument(
        "--output-dir",
        default="magalu_downloads_organizados",
        help="Pasta de saida para os arquivos organizados.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Se a pasta de saida ja existir, remove o conteudo e recria.",
    )
    return parser.parse_args()


def load_catalog(path: Path) -> dict[str, dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    return {
        item["id"]: item
        for item in payload
        if isinstance(item, dict) and isinstance(item.get("id"), str)
    }


def extract_report_id(path: Path) -> str | None:
    match = ID_PATTERN.search(path.name)
    return match.group(1) if match else None


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


def parse_sheet_month(path: Path) -> str | None:
    rows = parse_sheet_rows(path)
    if len(rows) < 2 or not rows[1]:
        return None

    try:
        date_value = EXCEL_EPOCH + timedelta(days=float(rows[1][0]))
    except ValueError:
        return None

    return date_value.strftime("%Y-%m")


def parse_pdf_month(path: Path) -> str | None:
    text = path.read_bytes().decode("latin-1", errors="ignore")

    match = PDF_ASCII_DATE_PATTERN.search(text)
    if match:
        return f"{match.group(1)[:4]}-{match.group(1)[4:6]}"

    match = PDF_HEX_DATE_PATTERN.search(text)
    if not match:
        return None

    try:
        decoded = bytes.fromhex(match.group(1)).decode("latin-1", errors="ignore")
    except ValueError:
        return None

    decoded_match = re.search(r"D:(\d{14})", decoded)
    if not decoded_match:
        return None

    value = decoded_match.group(1)
    return f"{value[:4]}-{value[4:6]}"


def month_from_name_or_report(path: Path, report: dict[str, Any] | None) -> str | None:
    range_match = DATE_RANGE_PATTERN.search(path.stem)
    if range_match:
        return f"{range_match.group(1)}-{range_match.group(2)}"

    parameters = report.get("parameters") if isinstance(report, dict) else None
    if isinstance(parameters, dict):
        for key in ("start_date", "end_date"):
            value = parameters.get(key)
            if isinstance(value, str) and re.match(r"20\d{2}-\d{2}-\d{2}", value):
                return value[:7]

    return None


def build_service_invoice_month_map(
    source_dir: Path,
    catalog: dict[str, dict[str, Any]],
) -> dict[tuple[str, str], str]:
    month_map: dict[tuple[str, str], str] = {}

    for path in sorted(source_dir.iterdir()):
        if not path.is_file() or path.name in EXCLUDED_NAMES:
            continue

        report_id = extract_report_id(path)
        if not report_id:
            continue

        report = catalog.get(report_id)
        if not report or report.get("type") != "service_invoices":
            continue

        subtype = report.get("sub_type")
        external_id = report.get("parameters", {}).get("external_id")
        if not subtype or not external_id:
            continue

        month_label: str | None = None
        if path.suffix.lower() == ".xlsx":
            month_label = parse_sheet_month(path)
        elif path.suffix.lower() == ".pdf":
            month_label = parse_pdf_month(path)

        if month_label:
            month_map[(str(subtype), str(external_id))] = month_label

    for (subtype, external_id), month_label in list(month_map.items()):
        if subtype.endswith("_xlsx"):
            pdf_key = (subtype[:-5] + "_pdf", external_id)
            month_map.setdefault(pdf_key, month_label)
        elif subtype.endswith("_pdf"):
            xlsx_key = (subtype[:-4] + "_xlsx", external_id)
            month_map.setdefault(xlsx_key, month_label)

    return month_map


def base_name_for_file(path: Path) -> str:
    token_match = TOKEN_PATTERN.match(path.stem)
    if token_match:
        return token_match.group(1)

    report_id = extract_report_id(path)
    if report_id and path.stem.endswith("_" + report_id):
        return path.stem[: -(len(report_id) + 1)]

    return path.stem


def build_duplicate_group(
    path: Path,
    report: dict[str, Any] | None,
) -> str | None:
    if not report:
        return None

    ext = path.suffix.lower().lstrip(".")
    if report.get("type") == "service_invoices":
        subtype = report.get("sub_type")
        external_id = report.get("parameters", {}).get("external_id")
        if subtype and external_id:
            return f"service::{subtype}::{ext}::{external_id}"

    return None


def build_sort_key(path: Path, report: dict[str, Any] | None) -> tuple[Any, ...]:
    report_creation = ""
    if isinstance(report, dict):
        report_creation = str(report.get("creation_date") or "")
    return (report_creation, path.stat().st_mtime, path.name)


def collect_entries(
    source_dir: Path,
    catalog: dict[str, dict[str, Any]],
) -> list[FileEntry]:
    service_month_map = build_service_invoice_month_map(source_dir, catalog)
    entries: list[FileEntry] = []

    for path in sorted(source_dir.iterdir()):
        if not path.is_file() or path.name in EXCLUDED_NAMES:
            continue

        ext = path.suffix.lower()
        if ext not in {".pdf", ".xlsx", ".csv"}:
            continue

        report_id = extract_report_id(path)
        report = catalog.get(report_id or "")

        month_label: str | None = None
        tokenized_name = TOKEN_PATTERN.match(path.stem) is not None
        if report and report.get("type") == "service_invoices":
            subtype = report.get("sub_type")
            external_id = report.get("parameters", {}).get("external_id")
            if subtype and external_id:
                month_label = service_month_map.get((str(subtype), str(external_id)))

        if not month_label and tokenized_name:
            if ext == ".xlsx":
                month_label = parse_sheet_month(path)
            elif ext == ".pdf":
                month_label = parse_pdf_month(path)

        if not month_label:
            month_label = month_from_name_or_report(path, report)

        if not month_label:
            month_label = "sem_mes"

        entries.append(
            FileEntry(
                path=path,
                report_id=report_id,
                report=report,
                ext=ext,
                month_label=month_label,
                base_name=base_name_for_file(path),
                duplicate_group=build_duplicate_group(path, report),
                sort_key=build_sort_key(path, report),
                uses_month_suffix=tokenized_name,
            )
        )

    return entries


def mark_duplicates(entries: list[FileEntry]) -> None:
    grouped: dict[str, list[FileEntry]] = defaultdict(list)

    for entry in entries:
        if entry.duplicate_group:
            grouped[entry.duplicate_group].append(entry)

    for group in grouped.values():
        if len(group) < 2:
            continue

        group.sort(key=lambda item: item.sort_key, reverse=True)
        keeper = group[0]
        keeper.bucket = "manter"

        for duplicate in group[1:]:
            duplicate.bucket = "duplicados"


def assign_target_names(entries: list[FileEntry]) -> None:
    bucket_name_counts: dict[tuple[str, str], int] = defaultdict(int)
    final_name_counts: dict[tuple[str, str], int] = defaultdict(int)

    for entry in sorted(entries, key=lambda item: (item.bucket, item.base_name, item.month_label, item.path.name)):
        if entry.uses_month_suffix:
            stem = f"{entry.base_name}_{entry.month_label}"
        else:
            stem = entry.base_name

        if entry.bucket == "duplicados":
            key = (entry.bucket, stem + entry.ext)
            bucket_name_counts[key] += 1
            stem = f"{stem}_dup{bucket_name_counts[key]:02d}"

        final_name = stem + entry.ext
        final_key = (entry.bucket, final_name)
        final_name_counts[final_key] += 1
        if final_name_counts[final_key] > 1:
            stem = f"{stem}_{final_name_counts[final_key]:02d}"
            final_name = stem + entry.ext

        entry.target_name = final_name


def ensure_output_dir(path: Path, overwrite: bool) -> None:
    if path.exists() and overwrite:
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def copy_entries(output_dir: Path, entries: list[FileEntry]) -> list[dict[str, str]]:
    manifest_rows: list[dict[str, str]] = []

    for entry in entries:
        bucket_dir = output_dir / entry.bucket
        bucket_dir.mkdir(parents=True, exist_ok=True)
        destination = bucket_dir / entry.target_name
        shutil.copy2(entry.path, destination)

        manifest_rows.append(
            {
                "source": str(entry.path),
                "bucket": entry.bucket,
                "target": str(destination),
                "month": entry.month_label,
                "report_id": entry.report_id or "",
            }
        )

    return manifest_rows


def write_manifest(output_dir: Path, entries: list[FileEntry], manifest_rows: list[dict[str, str]]) -> None:
    summary = {
        "generated_at": datetime.now().isoformat(),
        "total_files": len(entries),
        "manter": sum(1 for item in entries if item.bucket == "manter"),
        "duplicados": sum(1 for item in entries if item.bucket == "duplicados"),
        "sem_mes": sum(1 for item in entries if item.month_label == "sem_mes"),
    }

    payload = {
        "summary": summary,
        "files": manifest_rows,
    }

    (output_dir / "manifest_organizado.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> int:
    args = parse_args()

    source_dir = Path(args.source_dir)
    output_dir = Path(args.output_dir)
    catalog = load_catalog(Path(args.catalog_json))

    ensure_output_dir(output_dir, overwrite=args.overwrite)

    entries = collect_entries(source_dir, catalog)
    mark_duplicates(entries)
    assign_target_names(entries)
    manifest_rows = copy_entries(output_dir, entries)
    write_manifest(output_dir, entries, manifest_rows)

    print(f"Arquivos copiados para: {output_dir}")
    print(f"Manter: {sum(1 for item in entries if item.bucket == 'manter')}")
    print(f"Duplicados: {sum(1 for item in entries if item.bucket == 'duplicados')}")
    print(f"Sem mes: {sum(1 for item in entries if item.month_label == 'sem_mes')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
