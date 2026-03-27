#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re, sys, unicodedata
from pathlib import Path
from typing import Dict, List
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    print("Erro: pdfplumber não está instalado. Instale com: pip install pdfplumber")
    sys.exit(1)

try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    print("Erro: PyPDF2 não está instalado. Instale com: pip install PyPDF2")
    sys.exit(1)


def simplify_name(name: str) -> str:
    """
    Normaliza o nome da empresa:
    - troca & por " E "
    - remove acentos
    - deixa maiúsculo
    - remove caracteres estranhos
    - compacta espaços
    """
    name = name.replace("&", " E ")
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    name = name.upper()
    name = re.sub(r"[^A-Z0-9 ]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def extract_first_line_pdfplumber(page) -> str:
    """
    Tenta extrair a primeira linha de texto (visualmente) da página:
    - pega todas as "words"
    - descobre o menor 'top'
    - junta tudo o que estiver na mesma linha (com uma tolerância)
    """
    words = page.extract_words(x_tolerance=1, y_tolerance=1, keep_blank_chars=False)
    if not words:
        return ""
    min_top = min(w.get("top", 0) for w in words)
    tol = 4
    top_line_words = [w for w in words if abs(w.get("top", 0) - min_top) <= tol]
    if not top_line_words:
        return ""
    top_line_words.sort(key=lambda w: (w.get("x0", 0), w.get("x1", 0)))
    return " ".join(w.get("text", "").strip() for w in top_line_words if w.get("text", "").strip()).strip()


def extract_company_from_page(page) -> str:
    """
    Heurística para tentar descobrir o nome da empresa na página:
    - Tenta usar a "primeira linha" (via extract_first_line_pdfplumber)
    - Se não der, pega a primeira linha com texto do extract_text()
    """
    candidate = re.sub(r"\s+", " ", (extract_first_line_pdfplumber(page) or "")).strip()
    if candidate:
        # Ex: "12345 NOME DA EMPRESA LTDA"
        if re.match(r"^\d{1,6}\s+[A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç0-9\.\-& ]+$", candidate):
            return candidate
        # Ou algo razoável como primeira linha de texto
        if re.match(r"^[A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç0-9\.\-& ]{3,}$", candidate):
            return candidate
    try:
        txt = page.extract_text() or ""
        for ln in txt.splitlines():
            s = ln.strip()
            if s:
                return s
    except Exception:
        pass
    return "DESCONHECIDO"


def split_pdf_by_company_top_left(input_pdf: Path, out_dir: Path, competencia: str) -> Path:
    """
    Lê o PDF de holerites e separa as páginas por empresa (usando o topo da página),
    gerando um PDF por empresa e depois um ZIP com todos.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(str(input_pdf))
    created_paths: List[Path] = []
    company_pages: Dict[str, List[int]] = {}

    # Mapeia cada página -> "empresa" (chave simplificada)
    with pdfplumber.open(str(input_pdf)) as pdf:
        for idx, page in enumerate(pdf.pages):
            company_raw = extract_company_from_page(page)
            key = simplify_name(company_raw) if company_raw else f"DESCONHECIDO_PAG_{idx+1}"
            company_pages.setdefault(key, []).append(idx)

    # Gera um PDF por empresa
    for simple_key, pages in company_pages.items():
        writer = PdfWriter()
        for p in pages:
            writer.add_page(reader.pages[p])
        out_path = out_dir / f"{simple_key} {competencia}.pdf"
        with open(out_path, "wb") as f:
            writer.write(f)
        created_paths.append(out_path)

    # Compacta tudo em um ZIP
    zip_path = out_dir / f"{input_pdf.stem}_empresas_{competencia}.zip"
    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zf:
        for p in sorted(created_paths):
            zf.write(p, arcname=p.name)
    return zip_path


def main():
    if len(sys.argv) < 2:
        print("Uso: python SEPARAR_HOLERITES.py <arquivo.pdf> [competencia=MMYYYY]")
        sys.exit(1)

    input_pdf = Path(sys.argv[1])
    if not input_pdf.exists():
        print(f"Arquivo não encontrado: {input_pdf}")
        sys.exit(1)

    # competência: MMYYYY (padrão = mês/ano atual)
    competencia = sys.argv[2] if len(sys.argv) >= 3 else datetime.now().strftime("%m%Y")

    # >>> NOVO: timestamp para evitar conflito e casar com o monitor
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Pasta de saída baseada no nome do PDF (sem extensão) + timestamp
    # Ex: HOLERITES_20251119_094500
    pdf_prefix = input_pdf.stem.upper()
    out_dir = input_pdf.parent / f"{pdf_prefix}_{timestamp}"

    print(f"Processando: {input_pdf}")
    print(f"Saída: {out_dir} (competência={competencia})")
    try:
        zip_path = split_pdf_by_company_top_left(input_pdf, out_dir, competencia)
        print(f"Concluído. ZIP: {zip_path}")
    except Exception as e:
        print(f"Erro: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
