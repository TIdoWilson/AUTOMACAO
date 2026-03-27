#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Divide um PDF de "Relatório de Férias" por empresa (nome no topo da página),
gera um PDF para cada empresa com nome simplificado ("EMPRESA COMPETENCIA.pdf")
e compacta tudo em um único .zip.
- Pergunta ao usuário: arquivo PDF de entrada, pasta de saída e a competência (ex.: 112025).
Requisitos: PyPDF2
    pip install PyPDF2
"""
import re
import sys
import io
import os
import unicodedata
from pathlib import Path
from typing import Dict, List
from zipfile import ZipFile, ZIP_DEFLATED

try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    print("Erro: PyPDF2 não está instalado. Instale com: pip install PyPDF2")
    sys.exit(1)

# --- Utilidades de interface ---
def pick_file_and_dir_and_comp():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, simpledialog
    except Exception as e:
        print("Não foi possível carregar tkinter. Use: pip install tk (ou rode no Python com Tk incluído).")
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
        "Informe a competência para os nomes dos arquivos (ex.: 112025):",
        initialvalue="112025",
    )
    if not competencia:
        messagebox.showinfo("Operação cancelada", "Nenhuma competência informada.")
        sys.exit(0)

    return Path(pdf_path), Path(out_dir), competencia.strip()

# --- Funções de processamento ---
def simplify_name(name: str) -> str:
    # Troca & por E para não perder semântica ao retirar símbolos
    name = name.replace("&", " E ")
    # Normaliza e remove acentos
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    # Maiúsculas
    name = name.upper()
    # Mantém apenas letras, dígitos e espaços
    name = re.sub(r"[^A-Z0-9 ]+", " ", name)
    # Colapsa espaços
    name = re.sub(r"\s+", " ", name).strip()
    return name

def extract_company_from_page_text(text: str) -> str:
    """
    Tenta extrair o nome da empresa da parte superior da página.
    Padrões utilizados:
    1) "<EMPRESA> Página: N"
    2) Linha imediatamente anterior à que contém "Folha de Pagamento"
    3) Primeira linha não vazia
    """
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    first_line = lines[0] if lines else ""
    # Padrão 1: "EMPRESA Página: N"
    m = re.search(r"(.+?)\s+P[aá]gina\s*:\s*\d+", first_line, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # Padrão 2: linha anterior a "Folha de Pagamento"
    for i, ln in enumerate(lines):
        if "Folha de Pagamento" in ln:
            if i > 0:
                return lines[i-1].strip()

    # Padrão 3: fallback para primeira linha
    return first_line.strip() or "DESCONHECIDO"

def split_pdf_by_company(input_pdf: Path, out_dir: Path, competencia: str) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(input_pdf))
    company_pages: Dict[str, List[int]] = {}

    # Associa cada página à sua empresa
    for idx, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception:
            # Alguns PDFs podem falhar na extração; ainda assim adicionamos a página a "DESCONHECIDO"
            text = ""
        company = extract_company_from_page_text(text) or f"DESCONHECIDO_PAG_{idx+1}"
        company_pages.setdefault(company, []).append(idx)

    created_files = []

    # Cria um PDF por empresa
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

    # Cria ZIP com todos os PDFs
    zip_name = f"{input_pdf.stem}_empresas_{competencia}.zip"
    zip_path = out_dir / zip_name
    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zf:
        for f in sorted(created_files):
            zf.write(f, arcname=f.name)

    return zip_path

def main():
    try:
        input_pdf, out_dir, competencia = pick_file_and_dir_and_comp()
        print(f"PDF de entrada: {input_pdf}")
        print(f"Pasta de saída: {out_dir}")
        print(f"Competência: {competencia}")
        zip_path = split_pdf_by_company(input_pdf, out_dir, competencia)
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
