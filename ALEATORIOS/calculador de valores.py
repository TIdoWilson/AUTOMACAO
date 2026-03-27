import os
import math
from typing import List, Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


def to_cents(x) -> Optional[int]:
    """Converte para centavos (int). Aceita '1.234,56' e '1234.56'."""
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return None
        if isinstance(x, str):
            s = x.strip()
            if not s:
                return None
            s = s.replace(".", "").replace(",", ".")
            x = float(s)
        val = float(x)
        return int(round(val * 100))
    except Exception:
        return None


def subset_sum_bitset_indices(values: List[int], target: int) -> Optional[List[int]]:
    """
    Subset-sum EXATO para valores positivos usando bitset.
    Retorna índices (0-based) de 'values' que somam exatamente 'target', ou None.
    Complexidade típica: O(n * target/word) e memória ~ target bits.
    """
    if target < 0:
        return None

    # bit i ligado => soma i é atingível
    reachable = 1  # soma 0
    limit_mask = (1 << (target + 1)) - 1

    prev_sum = [-1] * (target + 1)
    prev_idx = [-1] * (target + 1)

    for i, v in enumerate(values):
        if v is None or v <= 0 or v > target:
            continue

        shifted = (reachable << v) & limit_mask
        new_bits = shifted & ~reachable

        # registrar predecessores apenas para somas "novas"
        b = new_bits
        while b:
            lsb = b & -b
            s = lsb.bit_length() - 1
            prev_sum[s] = s - v
            prev_idx[s] = i
            b -= lsb

        reachable |= shifted

        if (reachable >> target) & 1:
            # reconstrói
            res = []
            s = target
            while s != 0:
                idx = prev_idx[s]
                if idx == -1:
                    return None
                res.append(idx)
                s = prev_sum[s]
            res.reverse()
            return res

    return None


def main():
    root = tk.Tk()
    root.withdraw()

    path = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not path:
        return

    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        messagebox.showerror("Erro", f"Não consegui abrir o Excel.\n\nDetalhe: {e}")
        return

    sheet = simpledialog.askstring(
        "Aba (Sheet)",
        "Qual aba usar?\n\nOpções:\n- " + "\n- ".join(xl.sheet_names) +
        "\n\n(Deixe em branco para a primeira)"
    )
    if not sheet:
        sheet = xl.sheet_names[0]
    if sheet not in xl.sheet_names:
        messagebox.showerror("Erro", f"Aba '{sheet}' não encontrada.")
        return

    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Erro", f"Não consegui ler a aba '{sheet}'.\n\nDetalhe: {e}")
        return

    if df.empty:
        messagebox.showerror("Erro", "A aba está vazia.")
        return

    cols = list(df.columns)
    col = simpledialog.askstring(
        "Coluna",
        "Qual coluna somar?\n\nDigite exatamente o nome.\n\nColunas:\n- " + "\n- ".join(map(str, cols)) +
        "\n\n(Deixe em branco para a primeira coluna numérica)"
    )

    if not col:
        numeric_cols = [c for c in cols if pd.api.types.is_numeric_dtype(df[c])]
        if not numeric_cols:
            messagebox.showerror("Erro", "Não achei coluna numérica. Informe o nome da coluna manualmente.")
            return
        col = numeric_cols[0]

    if col not in df.columns:
        messagebox.showerror("Erro", f"Coluna '{col}' não encontrada.")
        return

    target_str = simpledialog.askstring("Valor esperado", "Qual o valor final esperado? (ex.: 8612,29)")
    if not target_str:
        return

    target = to_cents(target_str)
    if target is None:
        messagebox.showerror("Erro", "Valor esperado inválido.")
        return

    # prepara valores (somente positivos)
    idx_map = []   # posição em values -> índice do df (0-based)
    values = []    # centavos

    for df_i, raw in enumerate(df[col].tolist()):
        c = to_cents(raw)
        if c is None:
            continue
        if c <= 0:
            continue
        if c > target:
            continue
        values.append(c)
        idx_map.append(df_i)

    if not values:
        messagebox.showerror("Erro", "Não encontrei valores positivos <= alvo na coluna.")
        return

    # subset sum exato
    chosen = subset_sum_bitset_indices(values, target)
    if chosen is None:
        messagebox.showwarning(
            "Não encontrado",
            "Não foi encontrada combinação EXATA para o valor esperado.\n\n"
            "Se existir valor negativo ou precisar permitir valores > alvo, me avise que ajusto o método."
        )
        return

    chosen_df_rows = [idx_map[i] for i in chosen]  # índices do df (0-based)
    excel_rows_1based = [r + 2 for r in chosen_df_rows]  # +2 (cabeçalho + base 1)

    out_df = df.loc[chosen_df_rows].copy()
    out_df.insert(0, "__linha_excel_1based__", excel_rows_1based)

    out_path = os.path.join(os.path.dirname(path), "resultado_combinacao.csv")
    out_df.to_csv(out_path, index=False, encoding="utf-8-sig")

    total = sum(to_cents(df.loc[r, col]) or 0 for r in chosen_df_rows)

    messagebox.showinfo(
        "Resultado (EXATO)",
        f"Aba: {sheet}\nColuna: {col}\n\n"
        f"Alvo: R$ {target/100:.2f}\n"
        f"Encontrado: R$ {total/100:.2f}\n"
        f"Linhas no Excel (1-based): {', '.join(map(str, excel_rows_1based))}\n\n"
        f"Exportei para:\n{out_path}"
    )


if __name__ == "__main__":
    main()
