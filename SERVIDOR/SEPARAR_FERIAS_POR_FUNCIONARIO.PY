#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Splitador de PDF de Férias
--------------------------------------------------
O script:
- Pode pedir para o usuário selecionar um PDF (modo manual);
- Ou receber o caminho do PDF como argumento (modo automático, ex: chamado pelo monitor);
- Lê o nome da EMPRESA no PDF;
- Cria uma pasta "<NOME_DO_PDF>_YYYYMMDD_HHMMSS" no mesmo diretório do PDF;
- Divide o PDF em blocos de 2 páginas (Recibo + Aviso) para cada funcionário;
- Salva cada bloco como "<NOME_COMPLETO_FUNCIONARIO>_YYYYMMDD_HHMMSS.pdf".
- Se já existir um arquivo com o mesmo nome, ele NÃO sobrescreve e enumera:
  "<nome>_YYYYMMDD_HHMMSS.pdf", "<nome>_YYYYMMDD_HHMMSS 2.pdf", etc.

Requisitos:
    pip install PyPDF2
"""

from __future__ import annotations
import os
import re
import sys
from pathlib import Path
from datetime import datetime

try:
    # PyPDF2 >= 3.x
    from PyPDF2 import PdfReader, PdfWriter
except Exception:
    print("Erro ao importar PyPDF2. Instale com: pip install PyPDF2")
    raise


# ---------------------------------------------------------------------------
# Seleção de arquivo (GUI) com Tkinter - usado apenas quando não há argumento
# ---------------------------------------------------------------------------
def selecionar_pdf() -> Path:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception:
        print("Tkinter não está disponível. Informe o caminho do PDF como argumento:")
        print("    python SEPARAR_FERIAS_POR_FUNCIONARIO.PY /caminho/para/FERIAS.pdf")
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


# ---------------------------------------------------------------------------
# Funções de extração de texto e detecção de empresa / funcionário
# ---------------------------------------------------------------------------
def extrair_texto(pagina) -> str:
    try:
        return pagina.extract_text() or ""
    except Exception:
        # Alguns PDFs podem falhar. Nesse caso, retornamos string vazia.
        return ""


def encontrar_empresa(texto: str) -> str | None:
    """
    Tenta encontrar a razão social da empresa.
    Procura padrões como:
        EMPRESA : <nome>
        Empresa : <nome>
    """
    padroes = [
        r"EMPRESA\s*:\s*(.+)",
        r"Empresa\s*:\s*(.+)",
    ]
    for padrao in padroes:
        m = re.search(padrao, texto, flags=re.IGNORECASE)
        if m:
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
        m = re.search(
            r"Ilmo\s*Sr\(a\)\.\s*(.+?)\s+Código\s*:",
            t,
            flags=re.IGNORECASE | re.DOTALL,
        )
        if m:
            nome = m.group(1).strip()
            nome = nome.replace("\n", " ").strip(" :-\u200b")
            if nome:
                return nome

    return None


# ---------------------------------------------------------------------------
# Utilitários para nomes de arquivos
# ---------------------------------------------------------------------------
def sanitizar_para_arquivo(nome: str) -> str:
    """
    Remove caracteres problemáticos para nomes de arquivos.
    """
    nome = re.sub(r'[<>:"/\\|?*\n\r\t]', " ", nome)
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


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------
def main(caminho_pdf: Path | None = None) -> None:
    """
    Se caminho_pdf for None, abre o seletor de arquivo (modo manual).
    Se um Path for informado (ex.: via sys.argv[1]), processa diretamente (modo automático).
    """
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

    # Timestamp único desta execução (mesma data/hora para pasta e arquivos)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # ----------------------------------------------------------------------
    # Pasta de saída:
    #   - baseada no NOME DO PDF (sem extensão), em maiúsculo
    #   - adiciona _YYYYMMDD_HHMMSS
    #
    # Exemplo:
    #   PDF:   FERIAS_NOVEMBRO.pdf
    #   Pasta: FERIAS_NOVEMBRO_20251119_184210
    #
    # Isso é importante para o MONITOR_DE_PASTAS localizar a pasta depois,
    # usando o prefixo do nome do PDF.
    # ----------------------------------------------------------------------
    pdf_prefix = caminho_pdf.stem.upper()
    pasta_saida = caminho_pdf.parent / f"{pdf_prefix}_{timestamp}"
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

        # Padrão de nome: "<NOME_FUNCIONARIO>_YYYYMMDD_HHMMSS"
        base_nome = f"{sanitizar_para_arquivo(nome)}_{timestamp}"
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
