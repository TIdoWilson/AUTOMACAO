#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv
import os
import re
import shutil
from datetime import date


def normalize_value(v: str) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r'[<>:"/\\|?*\n\r\t]', "_", s)
    return s


def open_csv_with_fallback(path: str):
    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
    last_err = None
    for enc in encodings:
        try:
            f = open(path, "r", encoding=enc, newline="")
            f.read(2048)   # valida decoding
            f.seek(0)
            return f, enc
        except UnicodeDecodeError as e:
            last_err = e
    raise last_err


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    csv_path = os.path.join(base_dir, "empresasnmrlista.csv")
    recado_path = os.path.join(base_dir, "recado.pdf")
    out_dir = os.path.join(base_dir, "saida")

    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Não encontrei {csv_path}")
    if not os.path.exists(recado_path):
        raise FileNotFoundError(f"Não encontrei {recado_path}")

    os.makedirs(out_dir, exist_ok=True)

    today = date.today()
    current_month = today.strftime("%m")
    current_year = today.strftime("%Y")
    ext = os.path.splitext(recado_path)[1]  # .pdf

    total = 0

    f, used_enc = open_csv_with_fallback(csv_path)
    with f:
        print(f"Lendo CSV com encoding: {used_enc}")
        reader = csv.reader(f)

        # ignora a primeira linha (cabeçalho)
        try:
            next(reader)
        except StopIteration:
            print("CSV vazio. Nada a fazer.")
            return

        for row in reader:
            if not row:
                continue

            key = normalize_value(row[0])
            if not key:
                continue

            base_name = f"{key} RECADO IMPORTANTE {current_month} {current_year}"
            out_path = os.path.join(out_dir, base_name + ext)

            if os.path.exists(out_path):
                i = 2
                while True:
                    alt = os.path.join(out_dir, f"{base_name} ({i}){ext}")
                    if not os.path.exists(alt):
                        out_path = alt
                        break
                    i += 1

            shutil.copy2(recado_path, out_path)
            total += 1
            print(f"Criado: {os.path.basename(out_path)}")

    print(f"\nConcluído. Total gerado: {total}")
    print(f"Pasta de saída: {out_dir}")


if __name__ == "__main__":
    main()
