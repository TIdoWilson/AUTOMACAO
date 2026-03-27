# conta_transitoria.py
import sys
import re
import glob
import argparse
from pathlib import Path
import pandas as pd
from itertools import combinations
from openpyxl import load_workbook  # usado no writer; leitura ignora estilos

# >>> importa a função do modo Excel assistido
from excel_assistido import open_excel_and_present_day

# ======= CONFIGURAÇÕES =======
COLUMN_ALIASES = {
    "date":   [r"^data(\s+do\s+lan[cç]amento)?$", r"^dt$", r"^emiss[aã]o$", r"^lan[cç]amento$", r"^data$"],
    "hist":   [r"^hist[oó]rico$", r"^hist[oó]rico\s+do\s+lan[cç]amento$", r"^historico$", r"^descri[cç][aã]o$", r"^descri[cç][aã]o\s+do\s+lan[cç]amento$"],
    "debit":  [r"^d[eé]bito$", r"^valor\s*d[eé]bito$", r"^vlr\s*d[eé]bito$", r"^debito$"],
    "credit": [r"^cr[eé]dito$", r"^valor\s*cr[eé]dito$", r"^vlr\s*cr[eé]dito$", r"^credito$"],
    "batch":  [r"^lote$", r"^n[ºo]\s*lote$"],
}

EPS = 0.01
USE_FIRST_NOTE_ONLY = True

# limites de busca (modo mensal)
MAX_CAND_FOR_KSUM = 160     # limite máximo de candidatos por dia para k-sum
MAX_FOR_ZERO_3SUM  = 140    # limite para procurar triplas que somam zero
MAX_FOR_ZERO_4SUM  = 90     # limite para procurar quadras que somam zero

# ======= LIMPEZA DO HISTÓRICO (porta da Macro) =======
CUT_AFTER_FIRST_SLASH = True
CUT_TRAILING_AFTER_HYPHEN = False

HIST_PREFIX_PATTERNS = [
    r"VENDA\s+DE\s+MERCADORIA\s+CF\.?NF\.?",
    r"VLR\.?\s*ICMS\s*ST\s*CF\.?NF\.?",
    r"DEVOLU[ÇC][AÃ]O\s*CF\.?NF\.?",
    r"VLR\s*BAIXA\s*DE\s*T[IÍ]TULO\s*CF\.?NR\.?",
    r"VLR\s*JUROS\s*E\s*MULTAS\s*S\/T[IÍ]TULO\s*NR\.?",
    r"BAIXA\s*DE\s*T[IÍ]TULO\s*CF\.?NR\.?",
    r"ENTRADA\s*DE\s*T[IÍ]TULO\s*POR\s*SUBSTITU[IÍ][CÇ][AÃ]O\s*CF\.?NF\.?",
    r"T[IÍ]TULOS?\s*A\s*PAGAR\s*CF\.?NR\.?",
    r"T[IÍ]TULO\s*A\s*RECEBER\s*CF\.?NR\.?",
    r"VLR\s*DESCONTOS?\s*CONCEDIDOS?\s*S\/T[IÍ]TULOS?\s*NR"
]

# ======= LEITURA COM DETECÇÃO DE CABEÇALHO =======
def _detect_header_row(temp_df: pd.DataFrame) -> int:
    header_row = None
    for i, row in temp_df.iterrows():
        row_str = row.astype(str).str.lower()
        has_date = row_str.str.contains(r"\bdata\b|lan[cç]amento|emiss[aã]o", regex=True, na=False).any()
        has_hist = row_str.str.contains(r"hist[oó]rico|descri[cç][aã]o", regex=True, na=False).any()
        has_deb  = row_str.str.contains(r"d[eé]bito|debito", regex=True, na=False).any()
        has_cred = row_str.str.contains(r"cr[eé]dito|credito", regex=True, na=False).any()
        if has_date and (has_hist or has_deb or has_cred) and (has_deb or has_cred):
            header_row = i
            break
    if header_row is None:
        for i, row in temp_df.iterrows():
            row_str = row.astype(str).str.lower()
            score = 0
            score += row_str.str.contains(r"\bdata\b|lan[cç]amento|emiss[aã]o", regex=True, na=False).any()
            score += row_str.str.contains(r"hist[oó]rico|descri[cç][aã]o", regex=True, na=False).any()
            score += row_str.str.contains(r"d[eé]bito|cr[eé]dito|debito|credito", regex=True, na=False).any()
            if score >= 2:
                header_row = i
                break
    if header_row is None:
        raise ValueError("Não consegui localizar a linha de cabeçalho com 'Data/Histórico/Débito/Crédito'.")
    return header_row

def _finalize_from_temp(temp: pd.DataFrame, header_row: int) -> pd.DataFrame:
    print(f"[INFO] Cabeçalho detectado na linha (0-based): {header_row}")
    header_vals = list(temp.iloc[header_row])
    cols = []
    for idx, c in enumerate(header_vals):
        if c is None:
            cols.append(f"Unnamed: {idx}")
        else:
            name = str(c).strip()
            cols.append(name if name and name.lower() != "nan" else f"Unnamed: {idx}")
    data_rows = temp.iloc[header_row + 1 : ].copy()
    data_rows.columns = cols
    df = data_rows.dropna(axis=1, how="all")
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def read_with_header_detection(path: Path) -> pd.DataFrame:
    """
    Reabre e re-salva com Excel/COM para limpar estilos problemáticos.
    Tolerante a ambientes onde não é permitido setar .Visible/.DisplayAlerts.
    """
    import tempfile, shutil
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        raise RuntimeError("Precisa de 'pywin32'. Instale com: pip install pywin32")

    tmpdir = Path(tempfile.mkdtemp(prefix="xls_clean_"))
    cleaned_path = tmpdir / (path.stem + "_clean.xlsx")

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")  # nova instância isolada
    # tente reduzir UI, mas não quebre se estiver bloqueado
    for attr, val in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(excel, attr, val)
        except Exception:
            pass

    try:
        # Se o arquivo abrir em "Protected View", podemos tentar Open com ReadOnly
        wb = excel.Workbooks.Open(str(path), ReadOnly=True)
        # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.SaveAs(str(cleaned_path), FileFormat=51)
        wb.Close(False)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass

    # agora lê a cópia limpa
    temp = pd.read_excel(cleaned_path, header=None)
    header_row = _detect_header_row(temp)
    df = _finalize_from_temp(temp, header_row)

    # limpa temporário
    try:
        shutil.rmtree(tmpdir)
    except Exception:
        pass

    return df

# ======= AUXILIARES =======
def normalize_cols(df: pd.DataFrame):
    colmap = {}
    for canon, patterns in COLUMN_ALIASES.items():
        for c in df.columns:
            name = str(c).strip().lower()
            if any(re.search(p, name) for p in patterns):
                colmap[canon] = c; break
    missing = [k for k in ("date", "hist", "debit", "credit") if k not in colmap]
    if missing:
        raise ValueError("Não consegui identificar as colunas obrigatórias: " + ", ".join(missing) +
                         f". Colunas disponíveis: {list(df.columns)}.\n→ Ajuste COLUMN_ALIASES.")
    print("[INFO] Colunas mapeadas:", {k: colmap[k] for k in colmap})
    return colmap

def to_number(x):
    if pd.isna(x): return pd.NA
    s = str(x).strip()
    if s == "": return pd.NA
    s = re.sub(r"[^\d,.\-]", "", s)
    if "," in s and "." in s and s.rfind(",") > s.rfind("."):
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return pd.NA

def parse_date(col):
    return pd.to_datetime(col, errors="coerce", dayfirst=True)

def consolidate_history(df, col_date, col_hist):
    df = df.copy()
    df["_has_date"] = ~df[col_date].isna()
    df["_hist_add"] = (~df["_has_date"]) & df[col_hist].notna() & (df[col_hist].astype(str).str.strip() != "")
    df["_anchor_idx"] = df.index.to_series().where(df["_has_date"]).ffill()
    cont = df[df["_hist_add"]]
    if not cont.empty:
        add_text = cont.groupby("_anchor_idx")[col_hist].apply(lambda s: " | ".join(s.astype(str)))
        for anchor, extra in add_text.items():
            base = str(df.at[anchor, col_hist]) if pd.notna(df.at[anchor, col_hist]) else ""
            sep = " | " if base and extra else ""
            df.at[anchor, col_hist] = f"{base}{sep}{extra}"
    df = df[df["_has_date"]].copy()
    df.drop(columns=["_has_date", "_hist_add", "_anchor_idx"], inplace=True)
    return df

# ======= LIMPEZA DO HISTÓRICO =======
def _strip_accents_lower(s: str) -> str:
    import unicodedata
    s_norm = unicodedata.normalize("NFD", s)
    s_norm = "".join(ch for ch in s_norm if unicodedata.category(ch) != "Mn")
    return s_norm.lower()

def clean_hist_text(text: str) -> str:
    if pd.isna(text):
        return text
    raw = str(text)
    norm = _strip_accents_lower(raw)
    for pat in HIST_PREFIX_PATTERNS:
        rx = re.compile(r"^\s*" + pat, re.IGNORECASE)
        m = rx.search(norm)
        if m:
            cut_len = len(m.group(0))
            raw = raw[cut_len:]
            norm = norm[cut_len:]
            break
    if CUT_AFTER_FIRST_SLASH:
        p = raw.find("/")
        if p >= 0:
            raw = raw[:p]
    if CUT_TRAILING_AFTER_HYPHEN:
        p = raw.find("-")
        if p >= 0:
            raw = raw[:p]
    raw = re.sub(r"\s+", " ", raw).strip()
    raw = re.sub(r"[.,;:]\s*$", "", raw)
    return raw

# ======= EXTRAÇÃO DE NOTA =======
NF_NUMBER = re.compile(r"(?:\bNF(?:E)?\b[^\d]{0,3})(\d{3,}(?:[.\s]\d{3})*(?:[\/\-]\d{1,3})?)", re.I)
GEN_NUMBER = re.compile(r"(?<!\d)(\d{3,}(?:[.\s]\d{3})*(?:[\/\-]\d{1,3})?)(?!\d)", re.I)
DATE_LIKE  = re.compile(r"\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b", re.I)

def _normalize_note(token: str) -> str:
    parts = re.split(r"([\/\-])", token, maxsplit=1)
    if len(parts) == 3:
        base, sep, suf = parts
    else:
        base, sep, suf = parts[0], "", None
    base = re.sub(r"[.\s]", "", base)
    if re.fullmatch(r"\d+", base):
        base = str(int(base))
    return f"{base}{sep}{suf}" if sep and suf is not None else base

def extract_note_ids(text):
    if pd.isna(text): return []
    raw = str(text)
    nf_hits = [m.group(1) for m in NF_NUMBER.finditer(raw)]
    nf_hits = [h for h in nf_hits if not DATE_LIKE.search(h)]
    if nf_hits:
        best = _normalize_note(nf_hits[-1])
        return [best]
    cands = []
    for m in GEN_NUMBER.finditer(raw):
        token = m.group(1)
        if DATE_LIKE.search(token):
            continue
        cands.append(_normalize_note(token))
    if not cands:
        return []
    cands.sort(key=lambda s: (("/" in s) or ("-" in s), len(re.sub(r"\D", "", s))), reverse=True)
    return [cands[0]] if USE_FIRST_NOTE_ONLY else cands

# ======= REDUÇÃO POR CANCELAMENTO (apenas no próprio dia) =======
def _to_cents(x: float) -> int:
    return int(round(float(x) * 100))

def reduce_day_pool(sample_df: pd.DataFrame) -> pd.DataFrame:
    df = sample_df.copy()
    df = df[df["Diferenca"].abs() > EPS].copy()
    if df.empty:
        return df

    keep = [True] * len(df)
    val_map = {}
    for i, v in enumerate(df["Diferenca"].astype(float).tolist()):
        c = _to_cents(v)
        val_map.setdefault(c, []).append(i)
    # remove pares v / -v no DIA (o argumento já é “por dia”)
    for c, idxs in list(val_map.items()):
        if c == 0 or not idxs:
            continue
        opp = -c
        if opp not in val_map:
            continue
        take = min(len(idxs), len(val_map[opp]))
        for _ in range(take):
            i = val_map[c].pop()
            j = val_map[opp].pop()
            keep[i] = False
            keep[j] = False
    df = df[keep].reset_index(drop=True)

    # 3-sum / 4-sum (somente se n pequeno)
    n = len(df)
    if n == 0:
        return df

    def remove_indices(idxs):
        nonlocal df
        mask = [True] * len(df)
        for i in idxs:
            mask[i] = False
        df = df[mask].reset_index(drop=True)

    if n <= MAX_FOR_ZERO_3SUM:
        vals = df["Diferenca"].astype(float).tolist()
        cents = [_to_cents(v) for v in vals]
        arr = sorted([(cents[i], i) for i in range(len(vals))], key=lambda t: t[0])
        m = len(arr)
        a = 0
        while a < m - 2:
            ca, ia = arr[a]
            l, r = a + 1, m - 1
            while l < r:
                cb, ib = arr[l]
                cc, ic = arr[r]
                s = ca + cb + cc
                if s == 0:
                    rem_idx = sorted([ia, ib, ic], reverse=True)
                    remove_indices(rem_idx)
                    return reduce_day_pool(df)
                elif s < 0:
                    l += 1
                else:
                    r -= 1
            a += 1

    n = len(df)
    if n <= MAX_FOR_ZERO_4SUM:
        vals = df["Diferenca"].astype(float).tolist()
        cents = [_to_cents(v) for v in vals]
        arr = sorted([(c, i) for i, c in enumerate(cents)], key=lambda t: t[0])
        m = len(arr)
        for a in range(m - 3):
            ca, ia = arr[a]
            for b in range(a + 1, m - 2):
                cb, ib = arr[b]
                l, r = b + 1, m - 1
                rem = - (ca + cb)
                while l < r:
                    cc, ic = arr[l]
                    cd, idd = arr[r]
                    s = cc + cd
                    if s == rem:
                        rem_idx = sorted([ia, ib, ic, idd], reverse=True)
                        remove_indices(rem_idx)
                        return reduce_day_pool(df)
                    elif s < rem:
                        l += 1
                    else:
                        r -= 1

    return df

# ======= SELEÇÃO 1..4 ITENS PARA BATER A DIFERENÇA DO DIA =======
def _close(a, b, eps=EPS):
    return abs(a - b) <= eps

def pick_responsible_sets(por_dia_nota, dias_com_diff):
    selected = {}
    for _, row in dias_com_diff.iterrows():
        dia = row["Dia"]
        target = float(row["Diferenca"])

        base = por_dia_nota[(por_dia_nota["Dia"] == dia) & (por_dia_nota["Diferenca"].abs() > EPS)].copy()
        if base.empty:
            selected[dia] = set()
            continue

        base = base.sort_values(["NotaID"]).reset_index(drop=True)

        reduced = reduce_day_pool(base[["NotaID", "Diferenca"]])
        if len(reduced) > MAX_CAND_FOR_KSUM:
            reduced = reduced.sort_values("Diferenca", key=lambda s: s.abs(), ascending=False).head(MAX_CAND_FOR_KSUM)

        diffs = list(zip(reduced["NotaID"], reduced["Diferenca"].astype(float)))
        found = set()

        for k in range(1, min(4, len(diffs)) + 1):
            solved = False
            for combo in combinations(diffs, k):
                s = sum(d for _, d in combo)
                if _close(s, target):
                    found = {n for n, _ in combo}
                    solved = True
                    break
            if solved:
                break

        if not found and diffs:
            diffs_sorted = sorted(diffs, key=lambda t: abs(t[1]), reverse=True)
            best = None; best_gap = float("inf")
            for k in range(1, min(4, len(diffs_sorted)) + 1):
                s = sum(d for _, d in diffs_sorted[:k])
                gap = abs(s - target)
                if gap < best_gap:
                    best_gap, best = gap, {n for n, _ in diffs_sorted[:k]}
            found = best or set()

        selected[dia] = found
    return selected

# ======= PIPELINE PRINCIPAL =======
def process_file(xlsx_path: Path, only_day=None):
    df_raw = read_with_header_detection(xlsx_path)
    if df_raw.empty:
        raise ValueError("Planilha vazia após detecção de cabeçalho.")

    colmap = normalize_cols(df_raw)
    col_date  = colmap["date"]
    col_hist  = colmap["hist"]
    col_deb   = colmap["debit"]
    col_cred  = colmap["credit"]
    col_batch = colmap.get("batch", None)

    df = consolidate_history(df_raw, col_date, col_hist)

    # limpeza do histórico (porta macro)
    df[col_hist] = df[col_hist].apply(clean_hist_text)

    df[col_date] = parse_date(df[col_date])
    df["Debito"] = df[col_deb].apply(to_number).astype("Float64")
    df["Credito"] = df[col_cred].apply(to_number).astype("Float64")
    df["Valor"] = (df["Debito"].fillna(0) - df["Credito"].fillna(0)).astype("Float64")

    # descarta colunas que não ajudam (se existirem)
    for useless in ("Lote", "Saldo", "Contrap.", "Contrapartida"):
        if useless in df.columns:
            df = df.drop(columns=[useless])

    df = df[df[col_date].notna()].copy()
    df = df[~(df["Debito"].fillna(0).eq(0) & df["Credito"].fillna(0).eq(0))].copy()

    # Dia
    df["Dia"] = df[col_date].dt.date

    # Modo focal
    if only_day is not None:
        df = df[df["Dia"] == only_day].copy()
        if df.empty:
            print(f"[INFO] Nenhum lançamento no dia {only_day}.")
            out_path = xlsx_path.with_name(xlsx_path.stem + "_relatorio.xlsx")
            with pd.ExcelWriter(out_path, engine="openpyxl") as xlw:
                pd.DataFrame({"Mensagem":[f"Sem dados no dia {only_day}"]}).to_excel(xlw, "Resumo_Mensal", index=False)
            print(f"Relatório salvo em: {out_path}")
            return
        # limites mais generosos
        global MAX_CAND_FOR_KSUM, MAX_FOR_ZERO_3SUM, MAX_FOR_ZERO_4SUM
        MAX_CAND_FOR_KSUM = max(MAX_CAND_FOR_KSUM, 400)
        MAX_FOR_ZERO_3SUM = max(MAX_FOR_ZERO_3SUM, 300)
        MAX_FOR_ZERO_4SUM = max(MAX_FOR_ZERO_4SUM, 160)

    # Nota (após limpeza do histórico)
    df["NotaIDs"] = df[col_hist].apply(extract_note_ids)
    df["NotaID"] = df["NotaIDs"].apply(lambda l: l[0] if l else "SEM_NOTA")

    # Resumo mensal (ou do dia)
    resumo_mensal = df.agg({"Debito": "sum", "Credito": "sum", "Valor": "sum"}).to_frame(name="Total").T
    resumo_mensal["Fechou"] = (resumo_mensal["Valor"].abs() <= EPS)

    # Totais por dia
    dias = df.groupby("Dia", as_index=False).agg(Debito=("Debito","sum"), Credito=("Credito","sum"))
    dias["Diferenca"] = (dias["Debito"] - dias["Credito"]).astype(float)
    dias["Fechou"] = dias["Diferenca"].abs() <= EPS
    dias_com_diff = dias[~dias["Fechou"]].sort_values("Dia")

    # Por nota no mês (informativo)
    por_nota_mes = df.groupby("NotaID", as_index=False).agg(Debito=("Debito","sum"), Credito=("Credito","sum"))
    por_nota_mes["Diferenca"] = (por_nota_mes["Debito"] - por_nota_mes["Credito"]).astype(float)
    por_nota_mes = por_nota_mes.sort_values(["Diferenca","NotaID"], ascending=[False, True])

    # Por dia + nota (para combinação) — agora SEM descartar pelo mês; o descarte é por dia
    por_dia_nota = df.groupby(["Dia","NotaID"], as_index=False).agg(
        Debito=("Debito","sum"),
        Credito=("Credito","sum"),
    )
    por_dia_nota["Diferenca"] = (por_dia_nota["Debito"] - por_dia_nota["Credito"]).astype(float)

    # Só dias problemáticos (e só notas com dif != 0)
    dias_problematicos = set(dias_com_diff["Dia"])
    diffs_por_nota = (
        por_dia_nota[
            (por_dia_nota["Dia"].isin(dias_problematicos)) &
            (por_dia_nota["Diferenca"].abs() > EPS)
        ].sort_values(["Dia","NotaID"])
    )

    # Seleciona notas responsáveis (redução + k-sum)
    selected_by_day = pick_responsible_sets(diffs_por_nota, dias_com_diff)

    # Marca “Provavel_Responsavel”
    diffs_por_nota["Provavel_Responsavel"] = diffs_por_nota.apply(
        lambda r: r["NotaID"] in selected_by_day.get(r["Dia"], set()), axis=1
    )

    # Linhas dos responsáveis
    chaves_dia_nota = {(d, n) for d, notas in selected_by_day.items() for n in notas}
    responsaveis = df[df.apply(lambda r: (r["Dia"], r["NotaID"]) in chaves_dia_nota, axis=1)].copy()

    # Dif da nota no dia (para a aba)
    diffs_key = diffs_por_nota[["Dia","NotaID","Diferenca"]].rename(columns={"Diferenca":"DiferencaNotaDia"})
    responsaveis = responsaveis.merge(diffs_key, on=["Dia","NotaID"], how="left")

    # ====== MODO EXCEL ASSISTIDO (somente dias ainda não resolvidos) ======
    # critérios: se para algum dia a soma das notas selecionadas não fecha exatamente a diferença do dia
    dias_assistir = []
    for _, row in dias_com_diff.iterrows():
        dia = row["Dia"]
        target = float(row["Diferenca"])
        notas_sel = selected_by_day.get(dia, set())
        if not notas_sel:
            dias_assistir.append(dia); continue
        soma_sel = diffs_por_nota.loc[
            (diffs_por_nota["Dia"] == dia) & (diffs_por_nota["NotaID"].isin(notas_sel)),
            "Diferenca"
        ].sum()
        if not _close(soma_sel, target):
            dias_assistir.append(dia)

    # executa Excel assistido nesses dias (um por vez; lento de propósito)
    for dia in dias_assistir:
        print(f"[ASSISTIDO] Abrindo Excel para o dia {dia}…")
        df_dia = df[df["Dia"] == dia].copy()
        target = float(dias.loc[dias["Dia"] == dia, "Diferenca"].iloc[0])
        try:
            open_excel_and_present_day(
                xlsx_path,
                df_day=df_dia,
                dia=dia
            )
        except Exception as e:
            print(f"[ASSISTIDO][ERRO] {dia}: {e}")

    # Organiza colunas
    show_cols = ["Dia", col_date, col_hist, "NotaID", "Debito", "Credito", "Valor", "DiferencaNotaDia"]
    if col_batch and col_batch in df.columns:
        # você pediu pra descartar; então não vamos incluir mesmo se existir
        pass
    responsaveis = responsaveis[show_cols].sort_values(["Dia","NotaID"])

    # Saída Excel
    out_path = xlsx_path.with_name(xlsx_path.stem + "_relatorio.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as xlw:
        resumo_mensal.to_excel(xlw, sheet_name="Resumo_Mensal", index=False)
        dias.to_excel(xlw, sheet_name="Totais_por_Dia", index=False)
        dias_com_diff.to_excel(xlw, sheet_name="Dias_com_Diferenca", index=False)
        por_nota_mes.to_excel(xlw, sheet_name="Notas_Mes", index=False)
        diffs_por_nota.to_excel(xlw, sheet_name="Diferencas_por_Nota", index=False)
        responsaveis.to_excel(xlw, sheet_name="Lancamentos_Responsaveis", index=False)

    # ======= Console =======
    print("\n== RESUMO MENSAL ==" if only_day is None else "\n== RESUMO DO DIA ==")
    print(resumo_mensal.to_string(index=False))
    if not dias_com_diff.empty:
        print("\n== DIAS COM DIFERENÇA ==")
        print(dias_com_diff.to_string(index=False))
        if only_day is not None:
            print(f"\n[INFO] Modo focal: exibindo e combinando somente o dia {only_day}.")
        def _fmt_brl(v):
            try:
                return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return str(v)
        print("\n== NOTAS SELECIONADAS COMO RESPONSÁVEIS ==")
        for d, notas in selected_by_day.items():
            if not notas:
                continue
            itens = []
            for n in sorted(notas):
                val = diffs_por_nota.loc[
                    (diffs_por_nota["Dia"] == d) & (diffs_por_nota["NotaID"] == n),
                    "Diferenca"
                ].sum()
                itens.append(f"{n} (R$ {_fmt_brl(val)})")
            print(f"{d} -> " + ", ".join(itens))
        if dias_assistir:
            dias_txt = ", ".join(str(d) for d in dias_assistir)
            print(f"\n[ASSISTIDO] Dias enviados ao Excel assistido: {dias_txt}")
    else:
        print("\nTodos os dias fecharam em R$ 0,00.")
    print(f"\nRelatório salvo em: {out_path}")

# ======= CLI =======
def _pick_file_dialog():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        path = filedialog.askopenfilename(title="Selecione a planilha do mês",
            filetypes=[("Planilhas Excel","*.xlsx *.xlsm *.xls")])
        return Path(path) if path else None
    except Exception:
        return None

def _fallback_latest_xlsx():
    here = Path(__file__).resolve().parent
    files = sorted((Path(p) for p in glob.glob(str(here / "*.xls*"))),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None

def _parse_cli():
    p = argparse.ArgumentParser()
    p.add_argument("xlsx", nargs="?", help="Caminho do arquivo Excel")
    p.add_argument("--dia", "--day", dest="only_day", help="Filtrar para um único dia (ex.: 26/03/2025 ou 2025-03-26)")
    return p.parse_args()

def _parse_only_day(s: str):
    try:
        d = pd.to_datetime(s, dayfirst=True, errors="raise").date()
        return d
    except Exception:
        print(f"[WARN] Não consegui interpretar o dia: {s}. Ignorando filtro.")
        return None

def main():
    args = _parse_cli()
    xlsx = Path(args.xlsx) if args.xlsx else None
    if xlsx is None: xlsx = _pick_file_dialog()
    if xlsx is None: xlsx = _fallback_latest_xlsx()
    if xlsx is None or not xlsx.exists():
        print("Não encontrei a planilha.\n→ Rode: python conta_transitoria.py \"CAMINHO\\arquivo.xlsx\" [--dia 26/03/2025]")
        sys.exit(3)
    only_day = _parse_only_day(args.only_day) if args.only_day else None
    print(f"[INFO] Processando: {xlsx}" + (f" | Dia focal: {only_day}" if only_day else ""))
    process_file(xlsx, only_day=only_day)

if __name__ == "__main__":
    main()
