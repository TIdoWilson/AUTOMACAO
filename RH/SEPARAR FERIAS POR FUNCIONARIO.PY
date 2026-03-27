#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Splitador de PDF de Férias
--------------------------------------------------
O script:
- Pede para o usuário selecionar um PDF;
- Lê o nome da EMPRESA no PDF;
- Cria uma pasta "ferias - <empresa>" no mesmo diretório do PDF original;
- Divide o PDF em blocos de 2 páginas (Recibo + Aviso) para cada funcionário;
- Salva cada bloco como "FERIAS - <nome completo do funcionario>.pdf".
- Se já existir um arquivo com o mesmo nome, ele NÃO sobrescreve e enumera:
  "FERIAS - <nome>.pdf", "FERIAS - <nome> 2.pdf", "FERIAS - <nome> 3.pdf", etc.

Requisitos:
    pip install PyPDF2
"""

from __future__ import annotations
import os
import re
import sys
from pathlib import Path

try:
    # PyPDF2 >= 3.x
    from PyPDF2 import PdfReader, PdfWriter
except Exception as e:
    print("Erro ao importar PyPDF2. Instale com: pip install PyPDF2")
    raise

# Seleção de arquivo (GUI) com Tkinter
def selecionar_pdf() -> Path:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception:
        print("Tkinter não está disponível. Informe o caminho do PDF como argumento:")
        print("    python split_ferias_pdf.py /caminho/para/FERIAS.pdf")
        sys.exit(1)

    root = tk.Tk()
    root.withdraw()
    root.update()
    file_path = filedialog.askopenfilename(
        title="Selecione o PDF de Férias",
        filetypes=[("Arquivos PDF", "*.pdf")],
    )
    root.update()
    if not file_path:
        messagebox.showinfo("Cancelado", "Nenhum arquivo selecionado.")
        sys.exit(0)
    return Path(file_path)


def extrair_texto(pagina) -> str:
    try:
        return pagina.extract_text() or ""
    except Exception:
        # Alguns PDFs podem falhar. Nesse caso, retornamos string vazia.
        return ""


def encontrar_empresa(texto: str) -> str | None:
    # Padrões comuns para "EMPRESA : <nome>"
    padroes = [
        r"EMPRESA\s*:\s*(.+)",
        r"Empresa\s*:\s*(.+)",
    ]
    for padrao in padroes:
        m = re.search(padrao, texto, flags=re.IGNORECASE)
        if m:
            # Pega só até o fim da linha
            empresa = m.group(1).strip()
            empresa = empresa.splitlines()[0].strip(" :-\u200b")
            return empresa
    return None


def encontrar_nome_funcionario(textos: list[str]) -> str | None:
    """
    Tenta achar o nome do funcionário nos textos das páginas do bloco.
    1) Linha "NOME COMPLETO : <nome>"
    2) No Aviso: linha com "Ilmo Sr(a)." seguida do nome antes de "Código:"
    """
    # 1) NOME COMPLETO :
    for t in textos:
        m = re.search(r"NOME\s+COMPLETO\s*:\s*(.+)", t, flags=re.IGNORECASE)
        if m:
            nome = m.group(1).strip()
            nome = nome.splitlines()[0].strip(" :-\u200b")
            if nome:
                return nome

    # 2) Ilmo Sr(a). <nome> Código:
    for t in textos:
        m = re.search(r"Ilmo\s*Sr\(a\)\.\s*(.+?)\s+Código\s*:", t, flags=re.IGNORECASE | re.DOTALL)
        if m:
            nome = m.group(1).strip()
            nome = nome.replace("\n", " ").strip(" :-\u200b")
            if nome:
                return nome

    return None


def sanitizar_para_arquivo(nome: str) -> str:
    # Remove caracteres problemáticos em nomes de arquivos no Windows/macOS/Linux
    nome = re.sub(r'[<>:"/\\|?*\n\r\t]', " ", nome)
    # Compacta espaços
    nome = re.sub(r"\s{2,}", " ", nome).strip()
    return nome


def caminho_unico(base_dir: Path, base_nome: str, ext: str = ".pdf") -> Path:
    """
    Gera um caminho único no formato:
    base_nome + ext
    base_nome + " 2" + ext
    base_nome + " 3" + ext
    ...
    """
    n = 1
    while True:
        if n == 1:
            candidato = base_dir / f"{base_nome}{ext}"
        else:
            candidato = base_dir / f"{base_nome} {n}{ext}"
        if not candidato.exists():
            return candidato
        n += 1


def main(caminho_pdf: Path | None = None) -> None:
    if caminho_pdf is None:
        caminho_pdf = selecionar_pdf()
    else:
        caminho_pdf = Path(caminho_pdf)

    if not caminho_pdf.exists():
        print(f"Arquivo não encontrado: {caminho_pdf}")
        sys.exit(1)

    reader = PdfReader(str(caminho_pdf))

    # Tenta detectar a empresa a partir das primeiras páginas (até 6, por segurança)
    empresa = None
    for idx in range(min(6, len(reader.pages))):
        texto = extrair_texto(reader.pages[idx])
        empresa = encontrar_empresa(texto)
        if empresa:
            break

    if not empresa:
        empresa = "sem_empresa"
        print("Aviso: não foi possível identificar a EMPRESA no PDF. Usando 'sem_empresa'.")
    else:
        print(f"Empresa detectada: {empresa}")

    # Pasta de saída
    pasta_saida = caminho_pdf.parent / f"FERIAS - {sanitizar_para_arquivo(empresa)}"
    pasta_saida.mkdir(exist_ok=True)

    total_paginas = len(reader.pages)
    if total_paginas % 2 != 0:
        print(f"Aviso: número ímpar de páginas ({total_paginas}). A última será ignorada.")
        total_blocos = total_paginas // 2
    else:
        total_blocos = total_paginas // 2

    print(f"Total de páginas: {total_paginas}")
    print(f"Total de funcionários (blocos de 2 páginas): {total_blocos}")

    for bloco in range(total_blocos):
        i = bloco * 2
        paginas_bloco = [reader.pages[i], reader.pages[i + 1]]
        textos_bloco = [extrair_texto(p) for p in paginas_bloco]

        nome = encontrar_nome_funcionario(textos_bloco)
        if not nome:
            nome = f"funcionario_{bloco+1:03d}"
            print(f"Aviso: não encontrei nome no bloco {bloco+1}. Usando '{nome}'.")
        else:
            print(f"Funcionário {bloco+1}: {nome}")

        # Padrão de nome: "FERIAS - <nome>"
        base_nome = f"FERIAS - {sanitizar_para_arquivo(nome)}"
        caminho_saida = caminho_unico(pasta_saida, base_nome, ext=".pdf")

        writer = PdfWriter()
        writer.add_page(paginas_bloco[0])
        writer.add_page(paginas_bloco[1])

        with open(caminho_saida, "wb") as f:
            writer.write(f)

    print(f"\nConcluído! Arquivos salvos em: {pasta_saida}")


if __name__ == "__main__":
    # Suporte a argumento via CLI também
    arg = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    main(arg)
