# -*- coding: utf-8 -*-
"""
Script para limpar relatórios bancários e padronizar colunas:
DATA (DD/MM/AAAA), HISTORICO (texto), DEBITO e CREDITO (R$ 1.000,00)

Funcionalidades:
- Pede arquivo ao usuário (CSV/XLS/XLSX)
- Detecta o cabeçalho verdadeiro mesmo se houver "lixo" antes
- Identifica/renomeia colunas mesmo com variações (acentos, plurais, etc.)
- Remove colunas de nº do lançamento e de saldo
- Concatena histórico que venha em múltiplas linhas
- Remove linhas sem valores (nem débito, nem crédito)
- Garante formatação BRL com vírgula para centavos e ponto para milhar
- Exporta CSV ou XLSX conforme opção escolhida

Autor: você :)
"""
import os
import re
import sys
import math
import time
import warnings
from decimal import Decimal, ROUND_HALF_UP
import unicodedata
from itertools import combinations
import pandas as pd
from difflib import SequenceMatcher

# GUI para escolher arquivo/saída
import tkinter as tk
from tkinter import filedialog, messagebox

# Alguns CSVs podem ter encoding variável
try:
    import chardet  # type: ignore
except ImportError:
    chardet = None

warnings.simplefilter("ignore", category=UserWarning)

# --------- Utilidades ---------
BRL_RE = re.compile(r'^\s*(?:R\$\s*)?[-(]?\d{1,3}(?:\.\d{3})*(?:,\d{2})?\)?\s*$')
DATE_CELL_RE = re.compile(r'^\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s*(?:\|\s*(.*))?$')
GENERIC_TOKENS = {"ref","nf","finance","fatura","nota","nfe","nr","num","numero","dpl","rcbto"," "}

def detect_encoding(path):
    if chardet is None:
        return "utf-8"
    with open(path, "rb") as f:
        data = f.read(200000)  # 200KB é suficiente para detectar
    res = chardet.detect(data)
    return res.get("encoding") or "utf-8"

def sniff_delimiter(path, encoding="utf-8"):
    import csv
    with open(path, "r", encoding=encoding, errors="ignore") as f:
        sample = f.read(4096)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;|\t")
        return dialect.delimiter
    except Exception:
        # fallback comum no BR
        return ";"
    
# --- Regra específica: RCBTO.DPL.<nota><parcela> → RCBTO.DPL.<nota> ---
RCBTO_DPL_PREFIX_RE = re.compile(r'^\s*(rcbto[\s\.\-]*dpl[\s\.\-]*)(\d+)', re.IGNORECASE)

def fix_rcbto_dpl_nota_parcela(text: str) -> str:
    """
    Se o HISTORICO começar com 'RCBTO.DPL.' + dígitos, considera o último dígito como nº da parcela
    e mantém apenas a parte da nota. Ex.: 'RCBTO.DPL.1201 ...' -> 'RCBTO.DPL.120 ...'
    """
    if not isinstance(text, str):
        return text
    m = RCBTO_DPL_PREFIX_RE.match(text)
    if not m:
        return text
    prefix, num = m.group(1), m.group(2)
    if len(num) >= 2:
        nota = num[:-1]   # corta a parcela (último dígito)
        start, end = m.span()
        return f"{prefix}{nota}{text[end:]}"
    return text
    
# HELPERS CONCILIACAO
STOPWORDS_HIST = {
    "nf","nota","dpl.","dpl","duplicata","vlr.","vlr","valor",
    "rcbto","recebimento","adto","adiantamento","pgto","pagamento"
}

def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s

def normalize_hist_basic(text: str) -> str:
    """Minúsculas, sem acentos, sem pontuação/esp., remove stopwords simples."""
    if not isinstance(text, str):
        text = str(text or "")
    t = strip_accents(text).lower()
    # remove caracteres especiais
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    tokens = [tok for tok in t.split() if tok and tok not in STOPWORDS_HIST]
    return "".join(tokens).strip()

def normalize_hist_relaxed(text: str) -> str:
    """Versão semelhante: tokeniza, remove stopwords e ordena tokens (token set)."""
    base = normalize_hist_basic(text)
    if not base:
        return ""
    toks = sorted(set(base.split()))
    return "".join(toks)

def normalize_hist_core(text: str) -> str:
    base = strip_accents(str(text or "")).lower()
    base = re.sub(r"[^a-z0-9\s]", " ", base)
    toks = []
    for t in base.split():
        if t.isdigit():       # joga fora tokens puramente numéricos (ex.: 6380)
            continue
        if t in GENERIC_TOKENS or t in STOPWORDS_HIST:
            continue
        if len(t) <= 1:
            continue
        toks.append(t)
    # token set ordenado
    return "".join(sorted(set(toks)))

def brl_to_decimal_series(col: pd.Series) -> pd.Series:
    return col.apply(lambda x: parse_brl(x) if pd.notna(x) else None)

def running_balance(df: pd.DataFrame) -> pd.Series:
    """
    Recebe DF com colunas DEBITO/CREDITO (strings 'R$ ...' ou Decimal)
    e devolve uma Series com o saldo acumulado linha a linha.
    """
    bal = Decimal(0)
    out = []
    for _, r in df.iterrows():
        deb = r.get("DEBITO")
        cre = r.get("CREDITO")

        # aceitar tanto Decimal quanto string formatada
        deb = deb if isinstance(deb, Decimal) else parse_brl(deb)
        cre = cre if isinstance(cre, Decimal) else parse_brl(cre)

        if deb is None:
            deb = Decimal(0)
        if cre is None:
            cre = Decimal(0)

        bal += (deb - cre)
        out.append(bal)
    return pd.Series(out, index=df.index)

def add_totals_row(df: pd.DataFrame) -> pd.DataFrame:
    # Totais (numéricos)
    deb_vals = df["DEBITO"].apply(lambda x: parse_brl(x) if isinstance(x, str) else x).fillna(Decimal(0))
    cre_vals = df["CREDITO"].apply(lambda x: parse_brl(x) if isinstance(x, str) else x).fillna(Decimal(0))
    saldo_vals = df["SALDO"].apply(lambda x: parse_brl(x) if isinstance(x, str) else x).fillna(Decimal(0))

    total_row = {
        "DATA": "",
        "HISTORICO": "TOTAL",
        "DEBITO": format_brl(deb_vals.sum()),
        "CREDITO": format_brl(cre_vals.sum()),
        "SALDO": format_brl(saldo_vals.iloc[-1] if len(saldo_vals) else Decimal(0)),
    }
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

# ----------------- RESOLVEDOR DE COMBINAÇÕES (com N×N) -----------------
def try_match_combinations(amounts_left, amounts_right, max_combo=3,
                           MAX_GROUP_NxN=7, COMBOS_LIMIT=20000, TIME_BUDGET_S=2.0):
    """
    amounts_*: lista [(idx, Decimal)]
    Estratégia: 1×1 → N×1 → 1×N → N×N (k,m >= 2..max_combo) com podas.
    Limites:
      - pula N×N se grupo > MAX_GROUP_NxN por lado
      - aborta busca quando ultrapassa COMBOS_LIMIT tentativas ou TIME_BUDGET_S segundos
    """
    t0 = time.perf_counter()
    def timed_out(cnt):
        return (cnt > COMBOS_LIMIT) or ((time.perf_counter() - t0) > TIME_BUDGET_S)

    # use inteiros em centavos para acelerar somas
    def to_cents_list(pairs):
        out = []
        for i, v in pairs:
            if v is None:
                continue
            cents = int((v * 100).to_integral_value())
            out.append((i, cents))
        return out

    L = to_cents_list(amounts_left)
    R = to_cents_list(amounts_right)

    matched_groups = []
    used_left = set(); used_right = set()
    attempts = 0

    # ---------- 1×1 (hash exato) ----------
    mapR = {}
    for j, vr in R:
        mapR.setdefault(vr, []).append(j)
    for i, vl in L:
        if vl in mapR and mapR[vl]:
            j = mapR[vl].pop(0)
            matched_groups.append(([i], [j]))
            used_left.add(i); used_right.add(j)

    Lrem = [(i, v) for i, v in L if i not in used_left]
    Rrem = [(j, v) for j, v in R if j not in used_right]
    if not Lrem or not Rrem:
        return matched_groups

    # ---------- N×1 ----------
    from itertools import combinations
    for k in range(2, max_combo+1):
        if not Lrem or not Rrem or timed_out(attempts):
            break
        mapR = {}
        for j, v in Rrem:
            mapR.setdefault(v, []).append(j)
        used_l_tmp = set(); used_r_tmp = set()
        idxL = [i for i, _ in Lrem]; vL = {i: v for i, v in Lrem}
        for comb in combinations(idxL, k):
            attempts += 1
            if timed_out(attempts): break
            s = sum(vL[i] for i in comb)
            if s in mapR and mapR[s] and not any(i in used_l_tmp for i in comb):
                j = mapR[s].pop(0)
                matched_groups.append((list(comb), [j]))
                used_l_tmp.update(comb); used_r_tmp.add(j)
        used_left.update(used_l_tmp); used_right.update(used_r_tmp)
        Lrem = [(i, v) for i, v in Lrem if i not in used_left]
        Rrem = [(j, v) for j, v in Rrem if j not in used_right]

    # ---------- 1×N ----------
    for k in range(2, max_combo+1):
        if not Lrem or not Rrem or timed_out(attempts):
            break
        mapL = {}
        for i, v in Lrem:
            mapL.setdefault(v, []).append(i)
        used_l_tmp = set(); used_r_tmp = set()
        idxR = [j for j, _ in Rrem]; vR = {j: v for j, v in Rrem}
        for comb in combinations(idxR, k):
            attempts += 1
            if timed_out(attempts): break
            s = sum(vR[j] for j in comb)
            if s in mapL and mapL[s] and not any(j in used_r_tmp for j in comb):
                i = mapL[s].pop(0)
                matched_groups.append(([i], list(comb)))
                used_l_tmp.add(i); used_r_tmp.update(comb)
        used_left.update(used_l_tmp); used_right.update(used_r_tmp)
        Lrem = [(i, v) for i, v in Lrem if i not in used_left]
        Rrem = [(j, v) for j, v in Rrem if j not in used_right]

    # ---------- N×N (com poda forte) ----------
    if Lrem and Rrem and (len(Lrem) <= MAX_GROUP_NxN) and (len(Rrem) <= MAX_GROUP_NxN) and not timed_out(attempts):
        Lrem = sorted(Lrem, key=lambda x: x[0])
        Rrem = sorted(Rrem, key=lambda x: x[0])

        # mapa de somas de R por tamanho
        sum_map_R = {}
        for m in range(2, max_combo+1):
            idxs = [j for j, _ in Rrem]; vR = {j: v for j, v in Rrem}
            for comb in combinations(idxs, m):
                attempts += 1
                if timed_out(attempts): break
                s = sum(vR[j] for j in comb)
                sum_map_R.setdefault((m, s), []).append(comb)
            if timed_out(attempts): break

        for k in range(2, max_combo+1):
            if timed_out(attempts): break
            idxs = [i for i, _ in Lrem if i not in used_left]; vL = {i: v for i, v in Lrem}
            if len(idxs) < k:
                continue
            used_l_tmp = set(); used_r_tmp = set()
            for combL in combinations(idxs, k):
                attempts += 1
                if timed_out(attempts): break
                if any(i in used_l_tmp for i in combL):
                    continue
                s = sum(vL[i] for i in combL)
                found = False
                for m in range(2, max_combo+1):
                    cand = sum_map_R.get((m, s), [])
                    for combR in cand:
                        if any(j in used_r_tmp or j in used_right for j in combR):
                            continue
                        matched_groups.append((list(combL), list(combR)))
                        used_l_tmp.update(combL); used_r_tmp.update(combR)
                        found = True
                        break
                    if found: break
            used_left.update(used_l_tmp); used_right.update(used_r_tmp)

    return matched_groups

COMPANY_SUFFIX = {"ltda", "me", "epp", "sa", "s/a", "s.a."}

def singularize_light(tok: str) -> str:
    # remove 's' final em palavras >=4 letras (ex.: financeiras -> financeira)
    if len(tok) >= 4 and tok.endswith("s"):
        return tok[:-1]
    return tok

PT_STOPWORDS = {
    # gramaticais bem comuns
    "de","da","do","das","dos","e","a","o","as","os","para","por","com","em",
    "no","na","nos","nas","ao","à","às","aos","pela","pelo","pelas","pelos",
    "sem","sob","sobre","entre","ate","até","contra"
}
# já existiam:
# STOPWORDS_HIST (nf, nota, dpl., duplicata, vlr., valor, rcbto, recebimento, adto, adiantamento, pgto, pagamento)
COMPANY_SUFFIX = {"ltda", "me", "epp", "sa", "s/a", "s.a."}

def singularize_light(tok: str) -> str:
    if len(tok) >= 4 and tok.endswith("s"):
        return tok[:-1]
    return tok

def normalize_hist_fuzzy(text: str) -> str:
    base = strip_accents(str(text or "")).lower()
    # separa dígitos colados em letras (ex.: 15111jota -> 15111 jota)
    base = re.sub(r"(?<=\d)(?=[a-z])|(?<=[a-z])(?=\d)", " ", base)
    base = re.sub(r"[^a-z0-9\s]", " ", base)

    toks = []
    for t in base.split():
        if not t:
            continue
        if any(ch.isdigit() for ch in t):    # joga fora tokens com dígitos
            continue
        if t in STOPWORDS_HIST or t in PT_STOPWORDS or t in COMPANY_SUFFIX:
            continue
        t = singularize_light(t)
        if len(t) == 1:
            continue
        toks.append(t)
    if not toks:
        return ""
    toks = sorted(set(toks))
    return "".join(toks)

def sim_ratio(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()

def token_set(s: str) -> set[str]:
    return set(s.split()) if s else set()

def jaccard(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / union if union else 0.0


# FINAL HELPERS CONCILIACAO
def normalize_colname(s: str) -> str:
    s = str(s or "").strip().lower()
    # remover acentos simples
    s = (s
         .replace("á","a").replace("à","a").replace("â","a").replace("ã","a")
         .replace("é","e").replace("ê","e")
         .replace("í","i")
         .replace("ó","o").replace("ô","o").replace("õ","o")
         .replace("ú","u").replace("ç","c"))
    s = re.sub(r'\s+', ' ', s)
    return s

def normalize_text_pt(s: str) -> str:
    s = str(s or "").strip().lower()
    s = (s
         .replace("á","a").replace("à","a").replace("â","a").replace("ã","a")
         .replace("é","e").replace("ê","e")
         .replace("í","i")
         .replace("ó","o").replace("ô","o").replace("õ","o")
         .replace("ú","u").replace("ç","c"))
    s = re.sub(r'\s+', ' ', s)
    return s

TOTAL_KEYS = [
    "total", "totais", "subtotal", "sub-total",
    "saldo", "saldo anterior", "saldo final",
    "resumo", "encerramento", "fechamento", "somatorio", "somatorio"
]

def is_totalizer_text(txt: str) -> bool:
    """
    Retorna True apenas se o texto parecer um totalizador (TOTAL, SALDO, SUBTOTAL etc.)
    e não um histórico normal que apenas menciona 'saldo' no meio da frase.
    """
    n = normalize_text_pt(txt)
    if not n:
        return False

    # casos clássicos de totalizadores
    patterns = [
        r"^total( geral| de| dos| credito| debito| créditos| débitos)?$",
        r"^subtotal",
        r"^sub[- ]?total",
        r"^saldo( anterior| final| atual| do dia| geral)?$",
        r"^resumo",
        r"^fechamento",
        r"^encerramento",
        r"^somatorio",
    ]
    for p in patterns:
        if re.search(p, n):
            return True

    # se o texto for muito curto (até 3 palavras) e contiver 'saldo' ou 'total', também é suspeito
    if len(n.split()) <= 3 and any(k in n for k in ["saldo", "total"]):
        return True

    # caso contrário, não é totalizador
    return False

def is_totalizer_row(hist: str, deb, cre) -> bool:
    """
    Detecta linha totalizadora baseada no texto e estrutura.
    Agora ignora casos em que 'saldo' aparece no meio do histórico.
    """
    if not isinstance(hist, str):
        hist = str(hist or "")
    if is_totalizer_text(hist):
        return True

    # Linhas com débito e crédito ao mesmo tempo continuam sendo totalizadoras
    if deb is not None and cre is not None:
        return True

    # Linhas curtas com 'total' ou 'saldo' no começo + valor único
    n = normalize_text_pt(hist)
    if (n.startswith("saldo") or n.startswith("total")) and (deb is not None or cre is not None):
        return True

    return False

def is_balance_col(name: str) -> bool:
    n = normalize_colname(name)
    return any(key in n for key in ["saldo", "balance"])

def is_launch_number_col(name: str) -> bool:
    n = normalize_colname(name)
    return any(key in n for key in ["lanc", "num", "nr", "nº", "no "]) and "documento" not in n and "agencia" not in n

def parse_brl(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    s = str(value).strip()
    if s == "" or s == "-":
        return None
    # Aceitar (1.234,56) como negativo
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    s = s.replace("R$", "").replace(" ", "")
    # Se usar vírgula como decimal e ponto como milhar
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        num = Decimal(s)
        if neg:
            num = -num
        return num
    except Exception:
        return None

def format_brl(value: Decimal | None) -> str:
    if value is None:
        return ""
    q = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    neg = q < 0
    q = abs(q)
    inteiro, frac = divmod(int(q * 100), 100)
    inteiro_str = f"{inteiro:,}".replace(",", ".")
    return f"-R$ {inteiro_str},{frac:02d}" if neg else f"R$ {inteiro_str},{frac:02d}"

def ensure_date_ddmmyyyy(s):
    if pd.isna(s):
        return None
    if isinstance(s, (pd.Timestamp, )):
        return s.strftime("%d/%m/%Y")
    txt = str(s).strip()
    # às vezes a célula traz "dd/mm/aaaa | texto"
    m = DATE_CELL_RE.match(txt)
    if m:
        from pandas import to_datetime
        dt = pd.to_datetime(m.group(1), dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.strftime("%d/%m/%Y")
    # tentativa geral
    dt = pd.to_datetime(txt, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return None
    return dt.strftime("%d/%m/%Y")

def read_any(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        # Lê primeira planilha
        return pd.read_excel(path, header=None, dtype=str)
    else:
        enc = detect_encoding(path)
        delim = sniff_delimiter(path, enc)
        return pd.read_csv(path, header=None, dtype=str, encoding=enc, sep=delim, engine="python", on_bad_lines="skip")

def find_header_row(df: pd.DataFrame, max_scan=20):
    """
    Tenta localizar a linha que contém o cabeçalho real.
    Procura por palavras-chave típicas nas primeiras linhas.
    """
    candidates = [
        "data", "dt",
        "historico", "hist", "descricao", "descricao do lancamento",
        "debito", "débito", "valor debito", "valor débito",
        "credito", "crédito", "valor credito", "valor crédito"
    ]
    for i in range(min(max_scan, len(df))):
        row_vals = [normalize_colname(x) for x in df.iloc[i].tolist()]
        score = sum(any(c in v for c in candidates) for v in row_vals if v)
        # precisa ter pelo menos 2 hits pra dar confiança
        if score >= 2:
            return i
    # fallback: primeira linha
    return 0

def pick_best_column_map(cols):
    """
    Mapeia colunas originais -> nomes padronizados
    """
    colmap = {}
    for c in cols:
        n = normalize_colname(c)
        if n == "":
            continue
        if any(x == n or n.startswith(x) for x in ["data", "dt"]):
            colmap[c] = "DATA"
        elif any(k in n for k in ["historico","hist","descricao","descrição","memo","observacao","observa"]):
            colmap[c] = "HISTORICO"
        elif any(k in n for k in ["debito","deb","débito","valor deb"]):
            colmap[c] = "DEBITO"
        elif any(k in n for k in ["credito","cred","crédito","valor cred"]):
            colmap[c] = "CREDITO"
        elif is_balance_col(c):
            colmap[c] = "__DROP_SALDO__"
        elif is_launch_number_col(c):
            colmap[c] = "__DROP_LANC__"
        else:
            # manter sem mapear por enquanto
            pass
    return colmap

def collapse_multiline_history(df):
    """
    Une linhas consecutivas de histórico quando:
    - A linha não tem DATA mas tem HISTORICO -> concatena à última linha "aberta"
    - OU a linha tem DATA repetida e só histórico (sem débitos/créditos) -> acumula
    Regras anti-bug:
    - Linhas totalizadoras (TOTAL/SALDO/etc. ou com débito e crédito juntos) são ignoradas.
    - Em continuação sem data, só preenche DEBITO/CREDITO se ainda estiverem vazios (não sobrescreve).
    """
    out_rows = []
    buffer = {"DATA": None, "HISTORICO": "", "DEBITO": None, "CREDITO": None}

    def flush():
        nonlocal buffer
        if buffer["DATA"] and (buffer["DEBITO"] is not None or buffer["CREDITO"] is not None):
            out_rows.append(buffer.copy())
        buffer = {"DATA": None, "HISTORICO": "", "DEBITO": None, "CREDITO": None}

    for _, r in df.iterrows():
        data = ensure_date_ddmmyyyy(r.get("DATA"))
        hist = (r.get("HISTORICO") or "").strip()
        deb = r.get("DEBITO")
        cre = r.get("CREDITO")

        deb = parse_brl(deb) if deb is not None else None
        cre = parse_brl(cre) if cre is not None else None

        # Às vezes DATA|HIST no mesmo campo de HISTORICO
        if not data and hist:
            m = DATE_CELL_RE.match(hist)
            if m:
                data = ensure_date_ddmmyyyy(m.group(1))
                hist = (m.group(2) or "").strip()

        # Filtrar totalizadores (antes de qualquer ação)
        if is_totalizer_row(hist, deb, cre):
            # Não mexe no buffer nem grava nada
            continue

        if data:
            # nova transação
            if buffer["DATA"] is not None:
                flush()
            buffer["DATA"] = data
            buffer["HISTORICO"] = hist
            if deb is not None:
                buffer["DEBITO"] = deb
            if cre is not None:
                buffer["CREDITO"] = cre
            continue

        # Continuação sem data -> concatena histórico e preenche valores APENAS se estiverem vazios
        if buffer["DATA"] is not None:
            if hist:
                if buffer["HISTORICO"]:
                    buffer["HISTORICO"] += " "
                buffer["HISTORICO"] += hist
            if deb is not None and buffer["DEBITO"] is None:
                buffer["DEBITO"] = deb
            if cre is not None and buffer["CREDITO"] is None:
                buffer["CREDITO"] = cre
        else:
            # lixo antes de iniciar primeira transação
            pass

    # flush final
    if buffer["DATA"] is not None:
        flush()

    return pd.DataFrame(out_rows, columns=["DATA","HISTORICO","DEBITO","CREDITO"])

def clean_report(df_raw: pd.DataFrame) -> pd.DataFrame:
    # 1) localizar linha de cabeçalho
    hdr = find_header_row(df_raw)
    df = df_raw.copy()
    df.columns = df.iloc[hdr].astype(str).tolist()
    df = df.iloc[hdr+1:].reset_index(drop=True)

    # 2) remover colunas totalmente vazias
    df = df.dropna(axis=1, how="all")

    # 3) mapear colunas
    colmap = pick_best_column_map(df.columns)
    # renomear as que conhecemos
    df = df.rename(columns=colmap)

    # 4) descartar saldos e nº de lançamento
    drop_cols = [c for c, std in colmap.items() if std in ("__DROP_SALDO__","__DROP_LANC__")]
    df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors="ignore")

    # 5) se faltar alguma coluna essencial, tentar heurísticas adicionais
    # (por exemplo, colunas que são 90% moeda -> pode ser debito/credito se nome for ambíguo)
    def likely_money_series(s: pd.Series) -> bool:
        sample = s.dropna().astype(str).head(50).tolist()
        if not sample:
            return False
        hits = 0
        for v in sample:
            v = v.strip()
            if v == "":
                continue
            if BRL_RE.match(v) or parse_brl(v) is not None:
                hits += 1
        return hits >= max(5, int(0.6*len(sample)))

    # nomes padronizados já existentes
    std_cols = set([colmap.get(c, "") for c in df.columns] + [c for c in df.columns if c in ("DATA","HISTORICO","DEBITO","CREDITO")])

    # criar cópias padronizadas se faltarem
    if "DATA" not in df.columns:
        # tente encontrar uma coluna que contenha datas
        for c in df.columns:
            vals = pd.to_datetime(df[c], dayfirst=True, errors="coerce")
            if vals.notna().mean() > 0.5:
                df["DATA"] = df[c]
                break

    if "HISTORICO" not in df.columns:
        # pegue a primeira coluna textual longa
        text_cols = sorted(df.columns, key=lambda c: -(df[c].astype(str).str.len().median()))
        for c in text_cols:
            if c in ("DEBITO","CREDITO","DATA"):
                continue
            if not likely_money_series(df[c]):
                df["HISTORICO"] = df[c]
                break

    # atribuir debito/credito se faltarem
    money_candidates = [c for c in df.columns if c not in ("DEBITO","CREDITO") and likely_money_series(df[c])]
    # heurística: se houver 2 colunas de dinheiro, a menor média negativa deve ser débito, a outra crédito
    if "DEBITO" not in df.columns or "CREDITO" not in df.columns:
        parsed = {c: df[c].apply(parse_brl) for c in money_candidates}
        # classifica pela soma dos negativos/positivos
        sums = {c: (pd.Series(v).dropna().sum() if len(pd.Series(v).dropna()) else Decimal(0)) for c, v in parsed.items()}
        # preferir nomes que contenham 'deb' ou 'cred'
        for c in money_candidates:
            n = normalize_colname(c)
            if "deb" in n and "DEBITO" not in df.columns:
                df["DEBITO"] = df[c]
            if "cred" in n and "CREDITO" not in df.columns:
                df["CREDITO"] = df[c]
        # se ainda faltou, use as duas primeiras
        if "DEBITO" not in df.columns and len(money_candidates) >= 1:
            df["DEBITO"] = df[money_candidates[0]]
        if "CREDITO" not in df.columns and len(money_candidates) >= 2:
            df["CREDITO"] = df[money_candidates[1]]

    # 6) manter somente as colunas relevantes
    keep = [c for c in ("DATA","HISTORICO","DEBITO","CREDITO") if c in df.columns]
    df = df[keep].copy()

    # 7) normalizações e colapso de histórico multiline
    # normalizar strings
    for c in df.columns:
        df[c] = df[c].astype(str).str.replace(r"\s+", " ", regex=True).str.strip().replace({"nan":"", "None":""})

    # colapsar
    df2 = collapse_multiline_history(df)

    # 8) remover linhas sem valores (nem débito nem crédito)
    mask_has_val = df2["DEBITO"].notna() | df2["CREDITO"].notna()
    df2 = df2[mask_has_val].reset_index(drop=True)

    # 9) formatar saída final
    df2["DATA"] = df2["DATA"].apply(lambda x: x if x else "")
    df2["HISTORICO"] = df2["HISTORICO"].fillna("").astype(str).str.strip()

    df2["DEBITO"] = df2["DEBITO"].apply(format_brl)
    df2["CREDITO"] = df2["CREDITO"].apply(format_brl)

    return df2

def choose_file():
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="Selecione o relatório (CSV/XLS/XLSX)",
        filetypes=[
            ("Planilhas Excel", "*.xlsx *.xls"),
            ("CSV", "*.csv"),
            ("Todos os arquivos", "*.*"),
        ]
    )
    return path

def choose_save(ext_default=".xlsx"):
    root = tk.Tk()
    root.withdraw()
    if ext_default == ".xlsx":
        ft = [("Planilha Excel", "*.xlsx"), ("CSV", "*.csv")]
        defname = "relatorio_limpo.xlsx"
    else:
        ft = [("CSV", "*.csv"), ("Planilha Excel", "*.xlsx")]
        defname = "relatorio_limpo.csv"
    out = filedialog.asksaveasfilename(
        title="Salvar arquivo limpo como...",
        defaultextension=ext_default,
        initialfile=defname,
        filetypes=ft
    )
    return out

# ===================== FUNÇÕES CONCILIAÇÃO =====================
from datetime import timedelta

EPS = Decimal("0.01")      # tolerância padrão de 1 centavo
DATE_WIN = 2               # janela de datas ±2 dias

def near_amount(a: Decimal, b: Decimal, eps: Decimal = EPS) -> bool:
    return abs(a - b) <= eps

def near_date(d1, d2, days: int = DATE_WIN) -> bool:
    if not d1 or not d2:
        return False
    # d1/d2 chegam como "dd/mm/aaaa" do df limpo; converter
    try:
        dt1 = pd.to_datetime(d1, dayfirst=True, errors="coerce")
        dt2 = pd.to_datetime(d2, dayfirst=True, errors="coerce")
    except Exception:
        return False
    if pd.isna(dt1) or pd.isna(dt2):
        return False
    return abs((dt1 - dt2).days) <= days

def reconcile_transactions(df_clean: pd.DataFrame, max_combo=3):
    """
    Recebe DF já limpo com colunas: DATA, HISTORICO, DEBITO (R$), CREDITO (R$)
    Retorna três DFs: 
      - conciliados (com SALDO e total),
      - a_conciliar (com SALDO e total),
      - conciliados_aprox (apenas linhas conciliadas por ±ε, com coluna OBS).
    """
    if df_clean.empty:
        empty = df_clean.copy()
        empty2 = df_clean.copy()
        empty3 = pd.DataFrame(columns=["DATA","HISTORICO","DEBITO","CREDITO","OBS"])
        return (empty, empty2, empty3)

    work = df_clean.copy()
    work["HISTORICO"] = work["HISTORICO"].apply(fix_rcbto_dpl_nota_parcela)
    # valores numéricos etc...
    work["_ORDER"] = pd.RangeIndex(start=0, stop=len(work), step=1)
    work["_DEB"] = brl_to_decimal_series(work["DEBITO"])
    work["_CRE"] = brl_to_decimal_series(work["CREDITO"])

    # normalizações de histórico (agora já com a nota corrigida)
    work["_HIST_STRICT"] = work["HISTORICO"].apply(normalize_hist_basic)
    work["_HIST_RELAX"]  = work["HISTORICO"].apply(normalize_hist_relaxed)
    work["_HIST_FUZZY"]  = work["HISTORICO"].apply(normalize_hist_fuzzy)
    work["_HIST_CORE"]   = work["HISTORICO"].apply(normalize_hist_core)


    # lado e valor positivo
    work["_side"] = work.apply(lambda r: "D" if (r["_DEB"] or Decimal(0)) > 0 else ("C" if (r["_CRE"] or Decimal(0)) > 0 else ""), axis=1)
    work["_amt"]  = work.apply(lambda r: r["_DEB"] if r["_side"]=="D" else (r["_CRE"] if r["_side"]=="C" else Decimal(0)), axis=1)

    matched_pairs = []  # lista de (idxs_debitos, idxs_creditos, chave_hist, modo)
    used = set()
    approx_registry = []  # registros de conciliação aproximada (±ε)

    # ---------- Passos auxiliares ----------
    def group_and_match(key_col, mode_name):
        nonlocal used, matched_pairs
        for key, sub in work.loc[~work.index.isin(used)].groupby(key_col):
            if not key:
                continue
            debs = [(i, a) for i, a in zip(sub.index, sub["_amt"]) if sub.loc[i, "_side"] == "D" and a is not None and a > 0]
            cres = [(i, a) for i, a in zip(sub.index, sub["_amt"]) if sub.loc[i, "_side"] == "C" and a is not None and a > 0]
            if not debs or not cres:
                continue
            groups = try_match_combinations(debs, cres, max_combo=max_combo)
            for gD, gC in groups:
                matched_pairs.append((gD, gC, key, mode_name))
                used.update(gD); used.update(gC)

    def group_and_match_fuzzy(amount_key="__amount__", hist_key="_HIST_FUZZY", threshold=0.8):
        """Casa 1×1 por mesmo valor e alta similaridade (SequenceMatcher)."""
        nonlocal used, matched_pairs
        remaining = work.loc[~work.index.isin(used)].copy()
        if remaining.empty:
            return
        remaining[amount_key] = remaining["_amt"]
        for amt, sub in remaining.groupby(amount_key):
            if amt is None or amt == Decimal(0):
                continue
            debs = sub[(sub["_side"] == "D")].copy()
            cres = sub[(sub["_side"] == "C")].copy()
            if debs.empty or cres.empty:
                continue
            debs = debs.sort_values("_ORDER"); cres = cres.sort_values("_ORDER")
            used_d = set(); used_c = set()
            for i, row_d in debs.iterrows():
                hs_d = row_d[hist_key]; best_j = None; best_s = 0.0
                for j, row_c in cres.iterrows():
                    if j in used_c: 
                        continue
                    hs_c = row_c[hist_key]
                    if not hs_d or not hs_c:
                        continue
                    s = sim_ratio(hs_d, hs_c)
                    if s > best_s:
                        best_s = s; best_j = j
                if best_j is not None and best_s >= threshold:
                    matched_pairs.append(([i], [best_j], f"{hist_key}:{amt}", "fuzzy"))
                    used_d.add(i); used_c.add(best_j)
            used.update(list(used_d)); used.update(list(used_c))

    def group_and_match_jaccard(amount_key="__amount__", hist_key="_HIST_FUZZY", threshold=0.55):
        nonlocal used, matched_pairs
        remaining = work.loc[~work.index.isin(used)].copy()
        if remaining.empty:
            return
        remaining[amount_key] = remaining["_amt"]
        remaining["_TOKSET"] = remaining[hist_key].apply(token_set)
        for amt, sub in remaining.groupby(amount_key):
            if amt is None or amt == Decimal(0):
                continue
            debs = sub[sub["_side"] == "D"].sort_values("_ORDER")
            cres = sub[sub["_side"] == "C"].sort_values("_ORDER")
            if debs.empty or cres.empty:
                continue
            used_d = set(); used_c = set()
            for i, row_d in debs.iterrows():
                best_j = None; best_s = 0.0
                t_d = row_d["_TOKSET"]
                if not t_d: 
                    continue
                for j, row_c in cres.iterrows():
                    if j in used_c: 
                        continue
                    t_c = row_c["_TOKSET"]
                    if not t_c:
                        continue
                    s = jaccard(t_d, t_c)
                    if s > best_s:
                        best_s = s; best_j = j
                if best_j is not None and best_s >= threshold:
                    matched_pairs.append(([i], [best_j], f"{hist_key}:{amt}", "jaccard"))
                    used_d.add(i); used_c.add(best_j)
            used.update(list(used_d)); used.update(list(used_c))

    # ---------- (1) Passo aproximado: ±ε + janela de datas ----------
    def group_and_match_approx(eps=EPS, date_win=DATE_WIN, hist_key="_HIST_RELAX"):
        """
        1x1 por valor aproximado (|Δ|<=eps) e data próxima (±date_win).
        Greedy: menor diferença e maior similaridade de tokens.
        Registra matches 'epsilon'; se Δ!=0, entra no approx_registry.
        """
        nonlocal used, matched_pairs, approx_registry
        remaining = work.loc[~work.index.isin(used)].copy()
        if remaining.empty:
            return
        remaining["_TOKSET"] = remaining[hist_key].apply(token_set)
        debs = remaining[remaining["_side"]=="D"].sort_values("_ORDER")
        cres = remaining[remaining["_side"]=="C"].sort_values("_ORDER")
        if debs.empty or cres.empty:
            return
        used_d = set(); used_c = set()
        for i, row_d in debs.iterrows():
            best = None  # (score_tuple, j_index, diff, jaccard)
            for j, row_c in cres.iterrows():
                if j in used_c:
                    continue
                if not near_date(row_d["DATA"], row_c["DATA"], days=date_win):
                    continue
                if not near_amount(row_d["_amt"], row_c["_amt"], eps=eps):
                    continue
                diff = row_d["_amt"] - row_c["_amt"]
                s_tok = jaccard(token_set(row_d[hist_key]), token_set(row_c[hist_key]))
                score = (abs(diff), -(s_tok or 0.0),)
                if best is None or score < best[0]:
                    best = (score, j, diff, s_tok)
            if best is not None:
                _, j, diff, s_tok = best
                matched_pairs.append(([i], [j], f"{hist_key}:approx", "epsilon"))
                used_d.add(i); used_c.add(j)
                if diff != 0:
                    approx_registry.append({
                        "idx_debito": i,
                        "idx_credito": j,
                        "diferenca_centavos": int((diff * 100).to_integral_value()),
                        "dif_abs_centavos": int((abs(diff) * 100).to_integral_value()),
                        "eps_centavos": int((eps * 100).to_integral_value()),
                    })
        used.update(list(used_d)); used.update(list(used_c))

    # ---------- (7) Duplicados/estornos (mirror) ----------
    ESTORNO_KEYS = ("ESTORNO","REVERS","CHARGEBACK","CANCEL","AJUSTE","DEVOL")
    def group_and_match_reversals(eps=EPS, date_win=max(DATE_WIN, 3), hist_key="_HIST_FUZZY"):
        nonlocal used, matched_pairs
        remaining = work.loc[~work.index.isin(used)].copy()
        if remaining.empty:
            return
        remaining["_TOKSET"] = remaining[hist_key].apply(token_set)
        debs = remaining[remaining["_side"]=="D"].sort_values("_ORDER")
        cres = remaining[remaining["_side"]=="C"].sort_values("_ORDER")
        if debs.empty or cres.empty:
            return
        used_d = set(); used_c = set()
        for i, row_d in debs.iterrows():
            txt_d = (row_d["HISTORICO"] or "").upper()
            has_kw = any(k in txt_d for k in ESTORNO_KEYS)
            best = None
            for j, row_c in cres.iterrows():
                if j in used_c:
                    continue
                if not near_date(row_d["DATA"], row_c["DATA"], days=date_win):
                    continue
                if not near_amount(row_d["_amt"], row_c["_amt"], eps=eps):
                    continue
                s_tok = jaccard(token_set(row_d[hist_key]), token_set(row_c[hist_key]))
                if not has_kw and s_tok < 0.5:
                    continue
                score = (abs(row_d["_amt"]-row_c["_amt"]), -(s_tok or 0.0))
                if best is None or score < best[0]:
                    best = (score, j, s_tok)
            if best is not None:
                _, j, s_tok = best
                matched_pairs.append(([i],[j], f"{hist_key}:reversal", "estorno/mirror"))
                used_d.add(i); used_c.add(j)
        used.update(list(used_d)); used.update(list(used_c))

    # ---------- (5) Parcelas/recorrências ----------
    RE_PARC = re.compile(r"\b(\d{1,2})/(\d{1,2})\b|PARC(?:EL|)\w*", re.IGNORECASE)
    def strip_parcela_tokens(text: str) -> str:
        if not text:
            return ""
        t = RE_PARC.sub(" ", text)
        return " ".join(t.split())

    def group_and_match_parcelas(mode_name="parcelas"):
        nonlocal used, matched_pairs
        remain = work.loc[~work.index.isin(used)].copy()
        if remain.empty:
            return
        mask_parc = (
            remain["_HIST_FUZZY"].astype(str).str.contains(RE_PARC, na=False)
            | remain["HISTORICO"].astype(str).str.contains(RE_PARC, na=False)
        )
        sub = remain[mask_parc].copy()
        if sub.empty:
            return
        sub["_NUCLEO"] = sub["_HIST_RELAX"].apply(strip_parcela_tokens)
        for key, g in sub.groupby("_NUCLEO"):
            if not key:
                continue
            debs = [(i, a) for i, a in zip(g.index, g["_amt"]) if g.loc[i, "_side"]=="D" and a>0]
            cres = [(i, a) for i, a in zip(g.index, g["_amt"]) if g.loc[i, "_side"]=="C" and a>0]
            if not debs or not cres:
                continue
            groups = try_match_combinations(debs, cres, max_combo=max_combo)
            for gD, gC in groups:
                matched_pairs.append((gD, gC, key, mode_name))
                used.update(gD); used.update(gC)

    # ===================== ORDEM DOS PASSES =====================
    # Passo 1: históricos idênticos
    group_and_match("_HIST_STRICT", "identico")

    # Passo 2: parcelas/recorrências
    group_and_match_parcelas("parcelas")

    # Passo 3: históricos semelhantes (token set) — multi↔multi via resolvedor
    group_and_match("_HIST_RELAX", "semelhante")

    # Passo 3b: semelhante "core" (ignora números e marcadores)
    group_and_match("_HIST_CORE", "semelhante_core")

    # Passo 4: fuzzy 1×1
    group_and_match_fuzzy(hist_key="_HIST_FUZZY", threshold=0.8)

    # Passo 5: jaccard 1×1
    group_and_match_jaccard(hist_key="_HIST_FUZZY", threshold=0.55)

    # Passo 0: aproximado (±ε + datas)
    group_and_match_approx(eps=EPS, date_win=DATE_WIN, hist_key="_HIST_RELAX")

    # Passo 0.5: reversals (espelhos/estornos)
    group_and_match_reversals(eps=EPS, date_win=max(DATE_WIN, 3), hist_key="_HIST_FUZZY")


    # -------------------- Montagem dos DFs --------------------
    matched_idx = sorted(list(used))
    conciliados = work.loc[matched_idx].copy().sort_values("_ORDER")
    a_conciliar = work.drop(index=matched_idx).copy().sort_values("_ORDER")

    # DF de aproximados (com OBS)
    def build_approx_df():
        if not approx_registry:
            return pd.DataFrame(columns=["DATA","HISTORICO","DEBITO","CREDITO","OBS"])
        rows = []
        for rec in approx_registry:
            i, j = rec["idx_debito"], rec["idx_credito"]
            dif = rec["diferenca_centavos"]
            obs = f"Conciliado com diferença de {abs(dif)} centavos (ε≤{rec['eps_centavos']})"
            for idx in [i, j]:
                r = work.loc[idx, ["DATA","HISTORICO","DEBITO","CREDITO"]].copy()
                rows.append({**r.to_dict(), "OBS": obs, "_ORDER": work.loc[idx, "_ORDER"]})
        out = pd.DataFrame(rows).sort_values("_ORDER").drop(columns=["_ORDER"])
        return out

    conciliados_aprox = build_approx_df()

    # Remover colunas auxiliares e formatar
    def finalize(df):
        if df.empty:
            return pd.DataFrame(columns=["DATA","HISTORICO","DEBITO","CREDITO","SALDO"])
        # garantir ordem original e índice limpo
        df = df.sort_values("_ORDER").copy()
        df = df[["DATA","HISTORICO","DEBITO","CREDITO","_ORDER"]].reset_index(drop=True)
        # saldo acumulado
        saldo_series = running_balance(df)
        df["SALDO"] = saldo_series.apply(format_brl)
        # retirar coluna técnica e adicionar total
        df = df.drop(columns=["_ORDER"])
        df = add_totals_row(df)
        return df

    return finalize(conciliados), finalize(a_conciliar), conciliados_aprox

# utilitário “salvar 2 abas” antigo (mantido p/ compatibilidade interna)
def choose_save_2tabs():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    out = filedialog.asksaveasfilename(
        title="Salvar Excel (XLSX) com 2 ou 3 abas",
        defaultextension=".xlsx",
        initialfile="relatorio_conciliado.xlsx",
        filetypes=[("Planilha Excel", "*.xlsx")]
    )
    return out

# FINAL FUNCOES CONCILIACAO

def main():
    try:
        # 1) Escolher arquivo de entrada
        in_path = choose_file()
        if not in_path:
            messagebox.showinfo("Cancelar", "Nenhum arquivo selecionado.")
            return

        # 2) Ler (CSV/XLS/XLSX) sem cabeçalho definido
        df_raw = read_any(in_path)
        if df_raw.empty:
            messagebox.showerror("Erro", "Arquivo vazio ou não pôde ser lido.")
            return

        # 3) Limpeza (detecta cabeçalho real, normaliza e colapsa histórico multiline)
        cleaned = clean_report(df_raw)

        # 4) Conciliação (idênticos/semelhantes; agora com ±ε, parcelas, reversals e N×N)
        conciliados_df, a_conciliar_df, aproximados_df = reconcile_transactions(cleaned, max_combo=5)

        # 5) Salvar SEMPRE como um único Excel; cria 3ª aba se houver “aproximados”
        out_path = choose_save_2tabs()
        if not out_path:
            messagebox.showinfo("Cancelar", "Nenhum local de salvamento selecionado.")
            return

        if not out_path.lower().endswith(".xlsx"):
            out_path = os.path.splitext(out_path)[0] + ".xlsx"

        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            conciliados_df.to_excel(writer, index=False, sheet_name="CONCILIADOS")
            a_conciliar_df.to_excel(writer, index=False, sheet_name="A CONCILIAR")
            if not aproximados_df.empty:
                aproximados_df.to_excel(writer, index=False, sheet_name="CONCILIADOS_APROX")

        msg_extra = "" if aproximados_df.empty else "\n(+ aba CONCILIADOS_APROX)"
        messagebox.showinfo("Concluído", f"Arquivo salvo em:\n{out_path}{msg_extra}")

    except Exception as e:
        messagebox.showerror("Erro inesperado", f"Ocorreu um erro:\n{e}")
        raise

if __name__ == "__main__":
    main()
