#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Splitar PDF por empresa (holerites/folhas) lendo o NOME DA EMPRESA no topo/esquerda de cada página,
mesmo quando o texto extraído perde o layout ao copiar/colar.

Como funciona:
  • Para cada página, usa pdfplumber para pegar as PALAVRAS com coordenadas.
  • Encontra a linha mais alta (menor 'top') e junta as palavras dessa linha (ordenadas por x).
  • Usa essa linha como "nome da empresa" (ex.: "163 FERRO E ACO LTDA").
  • Agrupa as páginas por nome (normalizado) e gera 1 PDF por empresa.
  • Pergunta PDF de entrada, pasta de saída e competência (ex.: 102025).
  • Cria também um único .zip com todos os PDFs.

Dependências:
    pip install pdfplumber PyPDF2
"""

import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List
from zipfile import ZipFile, ZIP_DEFLATED

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

# ---------------- UI helpers ----------------
def pick_file_dir_and_competencia():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, simpledialog
    except Exception:
        print("Falha ao carregar tkinter. Garanta um Python com Tk (ou instale 'tk').")
        raise

    root = tk.Tk()
    root.withdraw()
    root.update()

    pdf_path = filedialog.askopenfilename(
        title="Selecione o PDF de entrada",
        filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
    )
    if not pdf_path:
        messagebox.showinfo("Operação cancelada", "Nenhum arquivo selecionado.")
        sys.exit(0)

    out_dir = filedialog.askdirectory(title="Selecione a pasta de saída")
    if not out_dir:
        messagebox.showinfo("Operação cancelada", "Nenhuma pasta selecionada.")
        sys.exit(0)

    competencia = simpledialog.askstring(
        "Competência",
        "Informe a competência para os nomes dos arquivos (ex.: 102025):",
        initialvalue="102025",
    )
    if not competencia:
        messagebox.showinfo("Operação cancelada", "Nenhuma competência informada.")
        sys.exit(0)

    return Path(pdf_path), Path(out_dir), competencia.strip()

# ---------------- Utils ----------------
def simplify_name(name: str) -> str:
    name = name.replace("&", " E ")
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    name = name.upper()
    name = re.sub(r"[^A-Z0-9 ]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name

def extract_first_line_pdfplumber(page) -> str:
    words = page.extract_words(x_tolerance=1, y_tolerance=1, keep_blank_chars=False)
    if not words:
        return ""
    min_top = min(w.get("top", 0) for w in words)
    tol = 4
    top_line_words = [w for w in words if abs(w.get("top", 0) - min_top) <= tol]
    if not top_line_words:
        return ""
    top_line_words.sort(key=lambda w: (w.get("x0", 0), w.get("x1", 0)))
    line_text = " ".join(w.get("text", "").strip() for w in top_line_words if w.get("text", "").strip())
    return line_text.strip()

def extract_company_from_page(page) -> str:
    candidate = re.sub(r"\s+", " ", (extract_first_line_pdfplumber(page) or "")).strip()
    if candidate:
        # Padrão: número + nome (preferido)
        if re.match(r"^\d{1,6}\s+[A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç0-9\.\-& ]+$", candidate):
            return candidate
        # Ou só nome
        if re.match(r"^[A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç0-9\.\-& ]{3,}$", candidate):
            return candidate
    # Fallback: tenta texto corrido (pode vir em ordem diferente)
    try:
        txt = page.extract_text() or ""
        for ln in txt.splitlines():
            s = ln.strip()
            if s:
                return s
    except Exception:
        pass
    return "DESCONHECIDO"

# ---------------- Core ----------------
def split_pdf_by_company_top_left(input_pdf: Path, out_dir: Path, competencia: str) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(input_pdf))
    with pdfplumber.open(str(input_pdf)) as pdf:
        company_pages: Dict[str, List[int]] = {}
        for idx, page in enumerate(pdf.pages):
            company_raw = extract_company_from_page(page)
            key = simplify_name(company_raw) if company_raw else f"DESCONHECIDO_PAG_{idx+1}"
            company_pages.setdefault(key, []).append(idx)

    created_paths: List[Path] = []
    for simple_key, pages in company_pages.items():
        writer = PdfWriter()
        for p in pages:
            writer.add_page(reader.pages[p])
        out_path = out_dir / f"{simple_key} {competencia}.pdf"
        with open(out_path, "wb") as f:
            writer.write(f)
        created_paths.append(out_path)

    zip_path = out_dir / f"{input_pdf.stem}_empresas_{competencia}.zip"
    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zf:
        for p in sorted(created_paths):
            zf.write(p, arcname=p.name)

    return zip_path

def main():
    try:
        input_pdf, out_dir, competencia = pick_file_dir_and_competencia()
        print(f"PDF de entrada: {input_pdf}")
        print(f"Pasta de saída: {out_dir}")
        print(f"Competência: {competencia}")
        zip_path = split_pdf_by_company_top_left(input_pdf, out_dir, competencia)
        print("\nConcluído!")
        print(f"Arquivo ZIP gerado: {zip_path}")
    except KeyboardInterrupt:
        print("\nOperação cancelada pelo usuário.")
    except Exception as e:
        print("Ocorreu um erro inesperado:")
        print(e)
        sys.exit(1)

if __name__ == "__main__":
    main()
