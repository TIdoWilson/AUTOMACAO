#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Divide um PDF de "Relatório de Férias" por empresa (nome no topo da página),
gera um PDF para cada empresa e compacta tudo em um ZIP.
Versão adaptada para uso automático pelo monitor de pastas.
"""

import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import datetime

try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    print("Erro: PyPDF2 não está instalado. Instale com: pip install PyPDF2")
    sys.exit(1)


def simplify_name(name: str) -> str:
    name = name.replace("&", " E ")
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    name = name.upper()
    name = re.sub(r"[^A-Z0-9 ]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def extract_company_from_page_text(text: str) -> str:
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    first_line = lines[0] if lines else ""
    m = re.search(r"(.+?)\s+P[aá]gina\s*:\s*\d+", first_line, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()

    for i, ln in enumerate(lines):
        if "Folha de Pagamento" in ln:
            if i > 0:
                return lines[i-1].strip()

    return first_line.strip() or "DESCONHECIDO"


def split_pdf_by_company(input_pdf: Path, out_dir: Path, competencia: str) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(str(input_pdf))
    company_pages: Dict[str, List[int]] = {}

    for idx, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        company = extract_company_from_page_text(text) or f"DESCONHECIDO_PAG_{idx+1}"
        company_pages.setdefault(company, []).append(idx)

    created_files = []

    for company, pages in company_pages.items():
        writer = PdfWriter()
        for p in pages:
            writer.add_page(reader.pages[p])
        simple_name = simplify_name(company)
        filename = f"{simple_name} {competencia}.pdf"
        out_path = out_dir / filename
        with open(out_path, "wb") as f:
            writer.write(f)
        created_files.append(out_path)

    zip_name = f"{input_pdf.stem}_empresas_{competencia}.zip"
    zip_path = out_dir / zip_name
    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zf:
        for f in sorted(created_files):
            zf.write(f, arcname=f.name)
    return zip_path


def main():
    if len(sys.argv) < 2:
        print("Uso: python SEPARAR_RELATORIO_DE_FERIAS.py <arquivo.pdf>")
        sys.exit(1)

    input_pdf = Path(sys.argv[1])
    if not input_pdf.exists():
        print(f"Arquivo não encontrado: {input_pdf}")
        sys.exit(1)

    competencia = datetime.now().strftime("%m%Y")  # exemplo: 102025
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = input_pdf.parent / f"{input_pdf.stem.upper()}_{timestamp}"
    print(f"Processando: {input_pdf}")
    print(f"Saída: {out_dir}")

    try:
        zip_path = split_pdf_by_company(input_pdf, out_dir, competencia)
        print(f"Concluído com sucesso. ZIP gerado em: {zip_path}")
    except Exception as e:
        print(f"Erro: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
