# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import tkinter as tk
from tkinter import filedialog

from openpyxl import Workbook
from openpyxl.styles import Font


RECIBO_SPLIT_RE = re.compile(r"RECIBO DE ADIANTAMENTO DE F(?:E|.)RIAS")
NOME_RE = re.compile(r"NOME COMPLETO\s*\\tab\s*:\s*\\tab\s*([^\\\r\n]+)")
MESES_RE = re.compile(
    r"VALORES REFERENTES AO M(?:E|.)S\s*([A-Z]{3}/\d{4})\s*\\tab\s*VALORES REFERENTES AO M(?:E|.)S\s*([A-Z]{3}/\d{4})"
)
LIQUIDOS_RE = re.compile(
    r"TOTAL VALORES L(?:I|.)QUIDOS\s*\\tab\s*([\d\.,]+)\s*\\tab\s*([\d\.,]+)"
)


@dataclass
class RegistroFerias:
    nome_funcionario: str
    valor_ref_mes_x: str
    total_liquido_mes_x: float
    valor_ref_mes_y: str
    total_liquido_mes_y: float


def moeda_br_para_float(valor: str) -> float:
    return float(valor.replace(".", "").replace(",", "."))


def extrair_registros(texto_rtf: str) -> list[RegistroFerias]:
    registros: list[RegistroFerias] = []
    blocos = RECIBO_SPLIT_RE.split(texto_rtf)

    for bloco in blocos[1:]:
        nome_match = NOME_RE.search(bloco)
        meses_match = MESES_RE.search(bloco)
        liquidos_match = LIQUIDOS_RE.search(bloco)

        if not (nome_match and meses_match and liquidos_match):
            continue

        nome = nome_match.group(1).strip()
        mes_x = meses_match.group(1).strip().upper()
        mes_y = meses_match.group(2).strip().upper()
        liquido_x = moeda_br_para_float(liquidos_match.group(1).strip())
        liquido_y = moeda_br_para_float(liquidos_match.group(2).strip())

        registros.append(
            RegistroFerias(
                nome_funcionario=nome,
                valor_ref_mes_x=mes_x,
                total_liquido_mes_x=liquido_x,
                valor_ref_mes_y=mes_y,
                total_liquido_mes_y=liquido_y,
            )
        )

    return registros


def salvar_excel(registros: Iterable[RegistroFerias], destino: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo Ferias"

    cabecalho = [
        "NOME DE FUNCIONARIO",
        "VALOR REF MES X",
        "TOTAL VALORES LIQUIDOS MES X",
        "VALOR REF MES Y",
        "TOTAL VALORES LIQUIDOS MES Y",
    ]

    ws.append(cabecalho)

    for col in range(1, len(cabecalho) + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    for registro in registros:
        ws.append(
            [
                registro.nome_funcionario,
                registro.valor_ref_mes_x,
                registro.total_liquido_mes_x,
                registro.valor_ref_mes_y,
                registro.total_liquido_mes_y,
            ]
        )

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 30

    for row in ws.iter_rows(min_row=2, min_col=3, max_col=5):
        row[0].number_format = "#,##0.00"
        row[2].number_format = "#,##0.00"

    wb.save(destino)


def selecionar_arquivo_rtf() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo RTF",
        filetypes=[
            ("Arquivos RTF", "*.rtf"),
            ("Todos os arquivos", "*.*"),
        ],
    )
    root.destroy()
    if not arquivo:
        return None
    return Path(arquivo)


def main() -> None:
    arquivo_rtf = selecionar_arquivo_rtf()
    if arquivo_rtf is None:
        print("Nenhum arquivo selecionado.")
        return

    if not arquivo_rtf.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {arquivo_rtf}")

    saida: Path = arquivo_rtf.with_name(f"{arquivo_rtf.stem}_resumo_ferias.xlsx")

    texto = arquivo_rtf.read_text(encoding="latin-1", errors="ignore")
    registros = extrair_registros(texto)

    if not registros:
        raise ValueError("Nenhum registro valido foi encontrado no arquivo RTF.")

    salvar_excel(registros, saida)
    print(f"Arquivo gerado com sucesso: {saida}")
    print(f"Total de funcionarios processados: {len(registros)}")


if __name__ == "__main__":
    main()
