import os
import re
import sys
import datetime as dt
from dataclasses import dataclass
from typing import List, Optional, Dict, Any

import pandas as pd

# PDF text extraction
import pdfplumber

# Folder picker
import tkinter as tk
from tkinter import filedialog, messagebox

SITUACOES_ALVO = {"Vencido", "Parcial"}
SITUACOES_TODAS = {"Quitado", "Vencido", "Parcial", "A vencer", "Perdido"}

# regex helpers
RE_FUNC = re.compile(r"^Funcionário:\s*(.+?)(?:\s+Data de Admissão:|$)")
RE_EMPRESA = re.compile(r"^(.*?)(?:\s+Página:|$)")
RE_PERIODO_AQ = re.compile(r"(\d{2}/\d{2}/\d{4}\s+a\s+\d{2}/\d{2}/\d{4})")
RE_DATA = re.compile(r"\d{2}/\d{2}/\d{4}")
RE_SITUACAO = re.compile(r"\b(Quitado|Vencido|Parcial|A vencer|Perdido)\b")
RE_NUM = re.compile(r"^-?\d+(?:[.,]\d+)?$")
RE_EMPRESA_NO_ARQUIVO = re.compile(r"^\s*(\d+(?:-\d+)?)\s+RELATORIODEFERIAS\b", re.IGNORECASE)

def pick_folder() -> Optional[str]:
    root = tk.Tk()
    root.withdraw()
    root.update()
    folder = filedialog.askdirectory(title="Selecione a pasta com os RELATORIODEFERIAS")
    root.destroy()
    return folder or None

def normalize_number(token: str) -> Optional[float]:
    token = token.strip()
    if not token:
        return None
    # 12,5 -> 12.5
    token = token.replace(".", "").replace(",", ".") if "," in token else token
    try:
        return float(token)
    except ValueError:
        return None

def extract_company_name(lines: List[str]) -> str:
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        m = RE_EMPRESA.match(ln)
        if m:
            name = m.group(1).strip()
            return name if name else "N/D"
    return "N/D"

def extract_company_number_from_filename(pdf_path: str) -> str:
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    m = RE_EMPRESA_NO_ARQUIVO.match(stem)
    if m:
        return m.group(1)
    return "-"

def extract_limit_before_situacao(line: str, situacao_span: slice) -> str:
    """
    Pega o "Limite Concessão" como a última data antes da palavra de situação.
    Se não houver data, retorna '-'.
    """
    before = line[:situacao_span.start]
    dates = RE_DATA.findall(before)
    return dates[-1] if dates else "-"

def subtract_years_safe(base_date: dt.date, years: int) -> dt.date:
    target_year = base_date.year - years
    try:
        return base_date.replace(year=target_year)
    except ValueError:
        # Ajuste para datas como 29/02 em anos não bissextos.
        return base_date.replace(year=target_year, day=28)

def is_periodo_aquisitivo_older_than_years(periodo_aq: str, years: int = 5, ref_date: Optional[dt.date] = None) -> bool:
    m = re.match(r"^\s*(\d{2}/\d{2}/\d{4})\s+a\s+(\d{2}/\d{2}/\d{4})\s*$", periodo_aq or "")
    if not m:
        return False

    try:
        fim_periodo = dt.datetime.strptime(m.group(2), "%d/%m/%Y").date()
    except ValueError:
        return False

    hoje = ref_date or dt.date.today()
    corte = subtract_years_safe(hoje, years)
    return fim_periodo < corte

def to_float_pt(num_str: str):
    if num_str is None:
        return None
    s = num_str.strip()
    if not s:
        return None
    # remove coisas como 10*** -> 10
    m = re.search(r"-?\d+(?:[.,]\d+)?", s)
    if not m:
        return None
    s = m.group(0)
    # pt-BR
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

def extract_saldo_from_line(line: str) -> float | None:
    """
    Regras:
    - se após a situação vier "- - <saldo> ...", saldo = token[2]
    - caso contrário, saldo = token[2] (Tipo, Dias Gozados, Saldo, ...)
    """
    msit = RE_SITUACAO.search(line)
    if not msit:
        return None

    after = line[msit.end():].strip()
    if not after:
        return None

    tokens = after.split()

    # garante mínimo
    if len(tokens) < 3:
        return None

    # Caso típico: "Vencido - - 24 0 7" ou "A vencer - - 2,5 0 0"
    if tokens[0] == "-" and tokens[1] == "-":
        return to_float_pt(tokens[2])

    # Caso típico: "Parcial N 10 20 0 0" ou "Quitado N 30 0 0 0"
    # Layout esperado após situação: Tipo, Dias Gozados, Saldo, ...
    if len(tokens) >= 3:
        return to_float_pt(tokens[2])

    # fallback (se vier incompleto)
    return to_float_pt(tokens[-1])

def parse_pdf(pdf_path: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    all_lines: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            for ln in txt.splitlines():
                all_lines.append(ln.rstrip("\n"))

    empresa = extract_company_name(all_lines)
    numero_empresa = extract_company_number_from_filename(pdf_path)

    i = 0
    funcionario_atual = None
    while i < len(all_lines):
        line = all_lines[i].strip()

        # começo de funcionário
        mfunc = RE_FUNC.match(line)
        if mfunc:
            funcionario_atual = mfunc.group(1).strip()
            i += 1
            # ler linhas do funcionário até o próximo funcionário ou fim,
            # interrompendo ao encontrar "Quitado"
            while i < len(all_lines):
                ln = all_lines[i].strip()
                if RE_FUNC.match(ln):
                    # próximo funcionário
                    break

                msit = RE_SITUACAO.search(ln)
                if msit:
                    situacao = msit.group(1)
                    # Regra de negócio:
                    # - "A vencer" nunca entra na saída
                    # - ao encontrar "Quitado", considera as linhas abaixo deste funcionário como quitadas
                    if situacao == "A vencer":
                        i += 1
                        continue
                    if situacao == "Quitado":
                        break

                    if situacao in SITUACOES_ALVO:
                        periodo_aq = "-"
                        mp = RE_PERIODO_AQ.search(ln)
                        if mp:
                            periodo_aq = mp.group(1)

                        if is_periodo_aquisitivo_older_than_years(periodo_aq, years=5):
                            i += 1
                            continue

                        limit = extract_limit_before_situacao(ln, slice(msit.start(), msit.end()))
                        saldo = extract_saldo_from_line(ln)
                        if situacao == "Vencido" and saldo is not None and abs(saldo) < 1e-9:
                            i += 1
                            continue

                        rows.append(
                            {
                                "número da empresa": numero_empresa,
                                "empresa": empresa,
                                "funcionario": funcionario_atual,
                                "periodo aquisitivo": periodo_aq,
                                "limite de concessão": limit,
                                "situação": situacao,
                                "saldo": saldo,
                                "arquivo_origem": os.path.basename(pdf_path),
                            }
                        )

                i += 1

            # continua após bloco do funcionário (não consome a próxima linha de funcionário aqui)
            continue

        i += 1

    return rows

def main():
    folder = pick_folder()
    if not folder:
        print("Nenhuma pasta selecionada. Encerrando.")
        return

    pdfs = []
    for fn in os.listdir(folder):
        if "relatoriodeferias" in fn.lower() and fn.lower().endswith(".pdf"):
            pdfs.append(os.path.join(folder, fn))
    pdfs.sort()

    if not pdfs:
        messagebox.showinfo("Aviso", "Não encontrei PDFs com 'relatoriodeferias' no nome nessa pasta.")
        return

    all_rows: List[Dict[str, Any]] = []
    for p in pdfs:
        try:
            all_rows.extend(parse_pdf(p))
        except Exception as e:
            print(f"[ERRO] Falha ao processar {os.path.basename(p)}: {e}", file=sys.stderr)

    if not all_rows:
        messagebox.showinfo("Resultado", "Nenhuma linha com situação Vencido/Parcial encontrada (ou todas quitadas).")
        return

    df = pd.DataFrame(all_rows)

    df["razão social"] = df["empresa"]

    # Coluna "número da empresa" deve vir antes de "razão social"
    cols_main = ["número da empresa", "razão social", "funcionario", "periodo aquisitivo", "limite de concessão", "situação", "saldo"]
    df_main = df[cols_main].copy()

    # Filtra pela data de "limite de concessão" em janela de -90/+90 dias da data atual.
    hoje = dt.date.today()
    inicio_90_atras = hoje - dt.timedelta(days=90)
    fim_90_apos = hoje + dt.timedelta(days=90)

    df_main["_limite_dt"] = pd.to_datetime(df_main["limite de concessão"], format="%d/%m/%Y", errors="coerce")
    limite_date = df_main["_limite_dt"].dt.date

    mask_90_atras = limite_date.between(inicio_90_atras, hoje, inclusive="both")
    mask_90_apos = limite_date.between(hoje + dt.timedelta(days=1), fim_90_apos, inclusive="both")

    df_vencidos_90_atras = df_main.loc[mask_90_atras, cols_main].copy()
    df_vencendo_90_apos = df_main.loc[mask_90_apos, cols_main].copy()

    # ordenação opcional
    sort_cols = ["número da empresa", "razão social", "funcionario", "periodo aquisitivo", "situação"]
    df_vencidos_90_atras.sort_values(by=sort_cols, inplace=True, kind="stable")
    df_vencendo_90_apos.sort_values(by=sort_cols, inplace=True, kind="stable")

    out_path = os.path.join(folder, "ferias_pendentes_consolidado.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_vencidos_90_atras.to_excel(writer, index=False, sheet_name="Vencidos_90d_atras")
        df_vencendo_90_apos.to_excel(writer, index=False, sheet_name="Vencendo_90d_apos")

    messagebox.showinfo("Concluído", f"Excel gerado:\n{out_path}")
    print(f"OK: {out_path}")

if __name__ == "__main__":
    main()
