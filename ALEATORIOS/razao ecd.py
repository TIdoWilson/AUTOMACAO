# -*- coding: utf-8 -*-
"""
SPED ECD - Extrair lançamentos de uma conta (I200/I250) e exportar para Excel.

Fluxo:
1) Usuário seleciona arquivos .txt
2) Usuário informa a conta
3) Script encontra linhas I250 da conta, usando a data do último I200
4) Exporta para .xlsx

Compatível com arquivos UTF-8 e muito grandes (processamento linha a linha).
"""

from __future__ import annotations

import os
from decimal import Decimal, InvalidOperation
from typing import Iterable, Optional, Tuple

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell


def select_files() -> Tuple[str, ...]:
    """Abre seletor de arquivos. Se não conseguir (sem tkinter), pede no console."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        paths = filedialog.askopenfilenames(
            title="Selecione os arquivos SPED ECD (.txt)",
            filetypes=[("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        return tuple(paths)
    except Exception:
        raw = input("Informe os caminhos completos dos arquivos (separe por ;): ").strip()
        if not raw:
            return tuple()
        return tuple(p.strip().strip('"') for p in raw.split(";") if p.strip())


def ask_account() -> str:
    conta = input("Digite a conta para buscar (ex: 112102001): ").strip()
    return conta


def ddmmaaaa_to_ddmmyyyy(s: str) -> str:
    s = (s or "").strip()
    if len(s) == 8 and s.isdigit():
        return f"{s[0:2]}/{s[2:4]}/{s[4:8]}"
    return s  # fallback


def parse_decimal_ptbr(s: str) -> Optional[Decimal]:
    """Converte '9000', '9000,00', '9.000,00' para Decimal."""
    if s is None:
        return None
    t = s.strip()
    if not t:
        return None
    # Remove separador de milhar, troca decimal para ponto
    t = t.replace(".", "").replace(",", ".")
    try:
        return Decimal(t)
    except (InvalidOperation, ValueError):
        return None


def format_brl(v: Decimal) -> str:
    """Formata Decimal como 'R$ 9.000,00' (milhar '.' e centavos ',')."""
    v = v.quantize(Decimal("0.01"))
    sign = "-" if v < 0 else ""
    v = abs(v)

    inteiro = int(v)
    cent = int((v - Decimal(inteiro)) * 100)

    inteiro_fmt = f"{inteiro:,}".replace(",", ".")  # 9000 -> 9.000
    return f"{sign}R$ {inteiro_fmt},{cent:02d}"


def iter_sped_lines(filepath: str) -> Iterable[str]:
    # errors="replace" evita quebra se houver algum byte inválido
    with open(filepath, "r", encoding="utf-8", errors="replace", newline="") as f:
        for line in f:
            line = line.strip("\r\n")
            if line:
                yield line


def extract_rows(files: Iterable[str], conta_alvo: str):
    """
    Retorna linhas extraídas como tuplas:
    (data_ddmmyyyy, deb_cred, historico, valor_brl)
    """
    for path in files:
        current_date = ""
        current_lcto = ""  # opcional (não exportamos, mas pode usar se quiser)
        for line in iter_sped_lines(path):
            # Checagens rápidas sem split completo
            if line.startswith("|I200|"):
                parts = line.split("|")
                # Estrutura: | I200 | num_lcto | dt_lcto | vl_lcto | ... |
                # indices:   0   1        2         3       4
                current_lcto = parts[2].strip() if len(parts) > 2 else ""
                dt_raw = parts[3].strip() if len(parts) > 3 else ""
                current_date = ddmmaaaa_to_ddmmyyyy(dt_raw)

            elif line.startswith("|I250|"):
                parts = line.split("|")
                # Exemplo:
                # |I250|112102001||9000|C||0|Historico||
                # indices: 0  1     2        4   5      8
                conta = parts[2].strip() if len(parts) > 2 else ""
                if conta != conta_alvo:
                    continue

                valor_raw = parts[4].strip() if len(parts) > 4 else ""
                ind_dc = parts[5].strip().upper() if len(parts) > 5 else ""
                historico = parts[8].strip() if len(parts) > 8 else ""

                deb_cred = "Crédito" if ind_dc == "C" else ("Débito" if ind_dc == "D" else ind_dc)
                v = parse_decimal_ptbr(valor_raw)
                valor_fmt = format_brl(v) if v is not None else ""

                # Se não houver I200 antes, current_date pode ficar vazio
                yield (current_date, deb_cred, historico, valor_fmt)


def save_excel(rows: Iterable[Tuple[str, str, str, str]], output_path: str) -> None:
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Lancamentos")

    headers = ["Data", "Deb/Cred", "Histórico", "Valor"]
    header_cells = []
    for h in headers:
        c = WriteOnlyCell(ws, value=h)
        header_cells.append(c)
    ws.append(header_cells)

    for row in rows:
        ws.append(list(row))

    # Larguras fixas (write_only não permite auto-fit real)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 80
    ws.column_dimensions["D"].width = 16

    wb.save(output_path)


def ask_output_path(default_name: str) -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        out = filedialog.asksaveasfilename(
            title="Salvar Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel", "*.xlsx")]
        )
        return out or default_name
    except Exception:
        out = input(f"Caminho para salvar (ENTER para '{default_name}'): ").strip()
        return out or default_name


def main():
    files = select_files()
    if not files:
        print("Nenhum arquivo selecionado. Encerrando.")
        return

    conta = ask_account()
    if not conta:
        print("Conta não informada. Encerrando.")
        return

    default_name = f"lancamentos_{conta}.xlsx"
    output_path = ask_output_path(default_name)

    rows = extract_rows(files, conta)
    save_excel(rows, output_path)

    print(f"OK: Arquivo gerado em: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    main()
