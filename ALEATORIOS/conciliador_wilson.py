# conciliador_wilson.py
# Requisitos (para rodar em Python): pip install pdfplumber pandas openpyxl
# Depois você empacota em .exe com PyInstaller (se quiser, eu te passo o comando).

import os
import re
import unicodedata
from dataclasses import dataclass
from typing import Optional, List

import pdfplumber
import pandas as pd
from datetime import datetime
from difflib import SequenceMatcher
# tolerância para comparar valores (centavos)
VALOR_TOLERANCIA = 0.05

# =========================
# Utils
# =========================
def to_date(s: str) -> datetime:
    return datetime.strptime(s, "%d/%m/%Y")


def norm_txt(s: str) -> str:
    if s is None:
        return ""
    s = s.strip().upper()
    s = "".join(
        c for c in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(c)
    )
    s = re.sub(r"\s+", " ", s)
    return s


def name_score(a: str, b: str) -> float:
    a = norm_txt(a)
    b = norm_txt(b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def parse_brl_num(s: str) -> Optional[float]:
    if s is None:
        return None
    s = s.strip()
    if not s:
        return None
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def extract_text_lines(pdf_path: str) -> List[str]:
    lines: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            for ln in txt.splitlines():
                ln = ln.rstrip()
                if ln:
                    lines.append(ln)
    return lines

# =========================
# Modelos
# =========================
@dataclass
class TxRazao:
    data: str
    cliente: str
    historico: str
    debito: float
    credito: float
    valor: float
    raw: str

@dataclass
class TxFinanceiro:
    data: str
    cpf_cnpj: str
    cliente: str
    titulo: str
    titulo_base: str
    debito: float
    credito: float

# =========================
# Parsing RAZÃO (com limpeza corrigida)
# =========================
def parse_razao(pdf_path: str) -> pd.DataFrame:
    lines = extract_text_lines(pdf_path)

    date_re = re.compile(r"^\s*(\d{1,2}/\d{2}/\d{4})\s*-\s*(.*)$")
    money_re = re.compile(r"(\d{1,3}(?:\.\d{3})*,\d{2})")

    def clean_cliente(cliente_raw: str) -> str:
        s = cliente_raw

        # remove lixo comum
        s = re.sub(r"\bDT\s*NFISCAL:.*$", "", s, flags=re.IGNORECASE).strip()
        s = re.sub(r"\bNFISCAL:\s*\d{2}/\d{2}/\d{4}.*$", "", s, flags=re.IGNORECASE).strip()
        s = re.sub(r"\b\d{1,2}/\d{2}/\d{4}\b.*$", "", s).strip()

        # remove sufixo Revenda
        s = re.sub(r"\bREVENDA\s*:\s*.*$", "", s, flags=re.IGNORECASE).strip()

        # remove valores monetários (preserva sobrenome que aparece depois do valor na linha concatenada)
        s = money_re.sub(" ", s)

        # corta quando começar códigos numéricos grandes (contas/saldos)
        s = re.sub(r"\b\d{6,}\b.*$", "", s).strip()

        # remove marcadores soltos
        s = re.sub(r"\b[DC]\b", " ", s)
        s = re.sub(r"\bDT\b", " ", s, flags=re.IGNORECASE)

        s = re.sub(r"\s+", " ", s).strip()
        return s

    rows: List[TxRazao] = []
    buffer = ""

    def flush_buffer(buf: str):
        m = date_re.match(buf)
        if not m:
            return

        data = m.group(1)
        rest = m.group(2)

        mc = re.search(r"\bCLIENTE\s+(.+)$", rest, flags=re.IGNORECASE)
        cliente_raw = mc.group(1).strip() if mc else ""
        cliente = clean_cliente(cliente_raw)

        valores = money_re.findall(rest)
        valor_lcto = parse_brl_num(valores[0]) if valores else None
        if valor_lcto is None:
            return

        debito = float(valor_lcto)
        credito = 0.0

        rows.append(
            TxRazao(
                data=data,
                cliente=cliente,
                historico=rest,
                debito=debito,
                credito=credito,
                valor=debito - credito,
                raw=buf,
            )
        )

    for ln in lines:
        if date_re.match(ln):
            if buffer:
                flush_buffer(buffer)
            buffer = ln
        else:
            # ignora cabeçalhos
            if any(k in ln.upper() for k in ["RELATÓRIO:", "EMPRESA:", "PÁGINA:", "USUÁRIO:", "DT. LCTO.", "CONTA CONTÁBIL"]):
                continue
            if buffer:
                buffer = buffer + " " + ln

    if buffer:
        flush_buffer(buffer)

    df = pd.DataFrame([r.__dict__ for r in rows])
    df["cliente_norm"] = df["cliente"].map(norm_txt)
    df["valor_round"] = df["valor"].round(2)
    return df

# =========================
# Parsing FINANCEIRO
# =========================
def parse_financeiro(pdf_path: str) -> pd.DataFrame:
    lines = extract_text_lines(pdf_path)

    periodo_re = re.compile(r"PER[IÍ]ODO\s+(\d{2}/\d{2}/\d{4})\s+A\s+(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
    dia_re = re.compile(r"\bDIA\s*:\s*(\d{1,2})\b", re.IGNORECASE)

    # parcela pode ser "-1" ou "-01"
    det_re = re.compile(
        r"^\s*([0-9\./-]{11,18})\s+(.+?)\s+(\d+)\s+(-\d{1,2})\s+.*?(\d{1,3}(?:\.\d{3})*,\d{2})\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s*$"
    )

    dt_ini = None
    for ln in lines[:150]:
        mp = periodo_re.search(ln)
        if mp:
            dt_ini = mp.group(1)
            break

    if not dt_ini:
        ini_mm, ini_yyyy = "01", "1900"
    else:
        _, ini_mm, ini_yyyy = dt_ini.split("/")

    current_day = None
    out: List[TxFinanceiro] = []

    for ln in lines:
        md = dia_re.search(ln)
        if md:
            current_day = int(md.group(1))
            continue

        if current_day is None:
            continue

        m = det_re.match(ln)
        if not m:
            continue

        cpf_cnpj = m.group(1).strip()
        nome = m.group(2).strip()
        titulo_base = m.group(3).strip()
        parcela = m.group(4).strip()
        deb = parse_brl_num(m.group(5)) or 0.0
        cred = parse_brl_num(m.group(6)) or 0.0

        data = f"{current_day:02d}/{ini_mm}/{ini_yyyy}"
        titulo = f"{titulo_base} {parcela}"

        out.append(
            TxFinanceiro(
                data=data,
                cpf_cnpj=cpf_cnpj,
                cliente=nome,
                titulo=titulo,
                titulo_base=titulo_base,
                debito=float(deb),
                credito=float(cred),
            )
        )

    df = pd.DataFrame([t.__dict__ for t in out], columns=[
        "data", "cpf_cnpj", "cliente", "titulo", "titulo_base", "debito", "credito"
    ])

    if df.empty:
        amostra = "\n".join(lines[:40])
        raise RuntimeError(
            "Não capturei nenhuma linha do Financeiro. Regex não bateu.\n"
            "Amostra:\n" + amostra
        )

    df["cliente_norm"] = df["cliente"].map(norm_txt)
    df["valor"] = (df["debito"] - df["credito"]).round(2)
    return df

# =========================
# Conciliação
# =========================
def conciliar(df_razao: pd.DataFrame, df_fin: pd.DataFrame, valor_tol=0.05, dias_janela=31, limiar_nome=0.72):
    raz = df_razao.copy()
    raz["valor"] = raz["valor_round"].astype(float)
    raz["data_dt"] = raz["data"].map(to_date)

    fin_linhas = df_fin.copy()
    fin_linhas["data_dt"] = fin_linhas["data"].map(to_date)
    fin_linhas["valor"] = (fin_linhas["debito"] - fin_linhas["credito"]).round(2)

    fin_grupos = (
        fin_linhas.groupby(["data", "cpf_cnpj", "cliente_norm", "titulo_base"], as_index=False)
        .agg(
            cliente=("cliente", "first"),
            debito=("debito", "sum"),
            credito=("credito", "sum"),
            valor=("valor", "sum"),
            parcelas=("titulo", lambda s: ", ".join(sorted(set(s)))),
        )
    )
    fin_grupos["valor"] = fin_grupos["valor"].round(2)
    fin_grupos["data_dt"] = fin_grupos["data"].map(to_date)

    group_to_line_idxs = {}
    for idx, row in fin_linhas.iterrows():
        k = (row["data"], row["cpf_cnpj"], row["cliente_norm"], row["titulo_base"])
        group_to_line_idxs.setdefault(k, []).append(idx)

    fin_line_used = set()
    fin_group_used = set()
    raz_used = set()

    matched_rows = []
    match_count = 0

    def tipo_razao(hist: str) -> str:
        h = (hist or "").upper()
        if "LANCAMENTO DE CARTAO" in h or "LANÇAMENTO DE CARTÃO" in h:
            return "CARTAO"
        if "ADIANTAMENTO" in h:
            return "ADIANTAMENTO"
        if "VENDA A PRAZO" in h or "NF" in h:
            return "VENDA"
        return "OUTRO"

    def match_em_linhas(r):
        rv = round(float(r["valor"]), 2)
        candidates = fin_linhas[(fin_linhas["valor"] - rv).abs() <= valor_tol]
        if candidates.empty:
            return None

        best = None
        best_key = None
        for f_i, f in candidates.iterrows():
            if f_i in fin_line_used:
                continue
            diff_dias = abs((r["data_dt"] - f["data_dt"]).days)
            if diff_dias > dias_janela:
                continue
            sc = name_score(r["cliente"], f["cliente"])
            if sc < (limiar_nome - 0.05):
                continue
            diff_val = abs(float(f["valor"]) - rv)
            key = (diff_val, diff_dias, -sc)
            if best_key is None or key < best_key:
                best_key = key
                best = (f_i, f, sc, diff_val, diff_dias)
        return best

    def match_em_grupos(r):
        rv = round(float(r["valor"]), 2)
        candidates = fin_grupos[(fin_grupos["valor"] - rv).abs() <= valor_tol]
        if candidates.empty:
            return None

        best = None
        best_key = None
        for g_i, g in candidates.iterrows():
            if g_i in fin_group_used:
                continue
            diff_dias = abs((r["data_dt"] - g["data_dt"]).days)
            if diff_dias > dias_janela:
                continue
            sc = name_score(r["cliente"], g["cliente"])
            if sc < (limiar_nome - 0.10):
                continue
            diff_val = abs(float(g["valor"]) - rv)
            key = (diff_val, diff_dias, -sc)
            if best_key is None or key < best_key:
                best_key = key
                best = (g_i, g, sc, diff_val, diff_dias)
        return best

    for r_i, r in raz.iterrows():
        t = tipo_razao(r["historico"])

        if t in ("CARTAO", "ADIANTAMENTO"):
            picked = match_em_linhas(r)
            regra = "PARCELA (linha)"
            if picked is None:
                picked = match_em_grupos(r)
                regra = "GRUPO (fallback)"
        else:
            picked = match_em_grupos(r)
            regra = "GRUPO"
            if picked is None:
                picked = match_em_linhas(r)
                regra = "PARCELA (fallback)"

        if picked is None:
            continue

        match_count += 1
        mid = f"M{match_count:05d}"
        raz_used.add(r_i)

        if regra.startswith("PARCELA"):
            f_i, f, sc, diff_val, diff_dias = picked
            fin_line_used.add(f_i)
            matched_rows.append({
                "match_id": mid,
                "modo": "LINHA",
                "regra": regra,
                "data_razao": r["data"],
                "data_fin": f["data"],
                "cliente_razao": r["cliente"],
                "cliente_fin": f["cliente"],
                "cpf_cnpj": f["cpf_cnpj"],
                "titulo_base": f["titulo_base"],
                "parcelas": f["titulo"],
                "valor_razao": round(float(r["valor"]), 2),
                "valor_fin": float(f["valor"]),
                "diff_valor": round(diff_val, 2),
                "diff_dias": int(diff_dias),
                "score_nome": round(sc, 3),
                "historico_razao": r["historico"],
            })
        else:
            g_i, g, sc, diff_val, diff_dias = picked
            fin_group_used.add(g_i)

            gkey = (g["data"], g["cpf_cnpj"], g["cliente_norm"], g["titulo_base"])
            for li in group_to_line_idxs.get(gkey, []):
                fin_line_used.add(li)

            matched_rows.append({
                "match_id": mid,
                "modo": "GRUPO",
                "regra": regra,
                "data_razao": r["data"],
                "data_fin": g["data"],
                "cliente_razao": r["cliente"],
                "cliente_fin": g["cliente"],
                "cpf_cnpj": g["cpf_cnpj"],
                "titulo_base": g["titulo_base"],
                "parcelas": g["parcelas"],
                "valor_razao": round(float(r["valor"]), 2),
                "valor_fin": float(g["valor"]),
                "diff_valor": round(diff_val, 2),
                "diff_dias": int(diff_dias),
                "score_nome": round(sc, 3),
                "historico_razao": r["historico"],
            })

    casados = pd.DataFrame(matched_rows)
    so_razao = df_razao.drop(index=list(raz_used)).copy()
    so_fin = fin_linhas.drop(index=list(fin_line_used)).copy()

    if not casados.empty:
        casados = casados.sort_values(["modo", "diff_valor", "diff_dias", "score_nome"], ascending=[True, True, True, False])
    if not so_razao.empty:
        so_razao = so_razao.sort_values(["data", "cliente"])
    if not so_fin.empty:
        so_fin = so_fin.sort_values(["data", "cliente", "titulo_base", "titulo"])

    return casados, so_razao, so_fin

# =========================
# UI: Selecionar arquivos e mostrar resultado no final
# =========================
def pick_pdf(title: str) -> str:
    # tkinter é padrão do Windows; não precisa instalar.
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("PDF files", "*.pdf")],
    )
    if not path:
        messagebox.showerror("Cancelado", "Nenhum arquivo selecionado.")
        raise SystemExit(1)
    return path

def print_df(df: pd.DataFrame, cols: List[str], title: str, max_rows: int = 200):
    print("\n" + "=" * 80)
    print(title)
    print("=" * 80)
    if df.empty:
        print("(vazio)")
        return
    show = df.copy()
    for c in cols:
        if c not in show.columns:
            show[c] = ""
    show = show[cols].head(max_rows)
    print(show.to_string(index=False))
    if len(df) > max_rows:
        print(f"... ({len(df) - max_rows} linhas omitidas)")

def main():
    razao_pdf = pick_pdf("Selecione o PDF do RAZÃO (Cartão de Crédito)")
    financeiro_pdf = pick_pdf("Selecione o PDF do FINANCEIRO (Cartão de Crédito)")

    # salva xlsx na mesma pasta dos PDFs (usa a pasta do Razão como referência)
    out_dir = os.path.dirname(razao_pdf)
    out_xlsx = os.path.join(out_dir, "conciliacao_cartao.xlsx")

    df_razao = parse_razao(razao_pdf)
    df_fin = parse_financeiro(financeiro_pdf)

    casados, so_razao, so_fin = conciliar(
        df_razao,
        df_fin,
        valor_tol=VALOR_TOLERANCIA,
        dias_janela=31,
        limiar_nome=0.72,
    )

    # totais/delta
    total_razao = float(df_razao["valor_round"].astype(float).sum())
    total_fin = float((df_fin["debito"].astype(float) - df_fin["credito"].astype(float)).sum())
    delta = round(total_razao - total_fin, 2)

    # sobras e conferência
    soma_so_razao = round(float(so_razao["valor_round"].astype(float).sum()) if not so_razao.empty else 0.0, 2)
    soma_so_fin = round(float(so_fin["valor"].astype(float).sum()) if not so_fin.empty else 0.0, 2)
    fecha = round(soma_so_razao - soma_so_fin, 2)

    # salva excel na mesma pasta
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
        casados.to_excel(w, index=False, sheet_name="CASADOS")
        so_razao.to_excel(w, index=False, sheet_name="DIF_SO_RAZAO")
        so_fin.to_excel(w, index=False, sheet_name="DIF_SO_FIN")
        pd.DataFrame([{
            "total_razao": round(total_razao, 2),
            "total_financeiro": round(total_fin, 2),
            "delta_razao_menos_fin": delta,
            "soma_so_razao": soma_so_razao,
            "soma_so_fin": soma_so_fin,
            "fecha_(so_razao-so_fin)": fecha,
        }]).to_excel(w, index=False, sheet_name="RESUMO_DIF")

    # exibe na tela (console)
    print("\nOK:", out_xlsx)
    print("Casados:", len(casados))
    print("Só no Razão:", len(so_razao))
    print("Só no Financeiro:", len(so_fin))
    print("\nTOTAL_RAZAO:", round(total_razao, 2))
    print("TOTAL_FIN:", round(total_fin, 2))
    print("DELTA (R-F):", delta)
    print("SOMA_SO_RAZAO:", soma_so_razao)
    print("SOMA_SO_FIN:", soma_so_fin)
    print("FECHA (SO_RAZAO - SO_FIN):", fecha)

    # imprime os lançamentos "sem par" no final
    print_df(
        so_fin,
        cols=["data", "cpf_cnpj", "cliente", "titulo_base", "titulo", "valor"],
        title="LANÇAMENTOS SÓ NO FINANCEIRO (sem par no Razão)",
        max_rows=500,
    )
    print_df(
        so_razao,
        cols=["data", "cliente", "valor_round", "historico"],
        title="LANÇAMENTOS SÓ NO RAZÃO (sem par no Financeiro)",
        max_rows=500,
    )

    # pausa (para não fechar se rodar por duplo clique no exe)
    input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    main()
