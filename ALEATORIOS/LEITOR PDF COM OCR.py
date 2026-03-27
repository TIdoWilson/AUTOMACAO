# conversor_pdf_para_excel_multimodal_acelerado.py
# Python 3.12 — Tabula + pdfplumber + OCR (fallback aprimorado)

import os, re, sys, warnings
from pathlib import Path
import pandas as pd
import numpy as np
import time
from multiprocessing import Pool, cpu_count
from tqdm import tqdm

# ===================== CONFIG =====================
PASTA_PDFS  = r"C:\Users\Usuario\Pictures\Feedback"  # <- AJUSTE AQUI
NOME_EXCEL  = "extratos_consolidados.xlsx"
NOME_LOG    = "extratos_consolidados_log.txt"

# Se precisar apontar executáveis explicitamente:
POPPLER_BIN   = r""  # ex.: r"C:\Program Files\poppler-25.07.0\Library\bin"
TESSERACT_EXE = r""  # ex.: r"C:\Program Files\Tesseract-OCR\tesseract.exe"

COLUNAS_PADRAO = ["DATA", "HISTORICO", "DOCUMENTO", "VALOR", "LIBER", "SALDO"]
MAX_COLUNAS     = 20
MIN_COLUNAS_ACEITAS = 3
# ==================================================

if POPPLER_BIN and POPPLER_BIN not in os.environ.get("PATH", ""):
    os.environ["PATH"] = POPPLER_BIN + os.pathsep + os.environ["PATH"]

warnings.filterwarnings("ignore", category=UserWarning)

# ---------- helpers ----------
def make_unique(columns):
    seen, out = {}, []
    for c in map(str, columns):
        if c in seen:
            seen[c] += 1
            out.append(f"{c}.{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out

def limpar_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    df = df.fillna("").astype(str)
    df.replace({"\u00a0": " ", "\x0c": " "}, regex=True, inplace=True)
    # NÃO dropa linhas aqui!
    # Remove apenas colunas 100% vazias:
    empty_col = df.apply(lambda col: col.astype(str).str.strip().eq("").all())
    df = df.loc[:, ~empty_col]
    if df.shape[1] > MAX_COLUNAS:
        df = df.iloc[:, :MAX_COLUNAS]
    return df

def get_page_count(pdf_path: str) -> int:
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            return len(pdf.pages)
    except Exception:
        try:
            from PyPDF2 import PdfReader
            return len(PdfReader(pdf_path).pages)
        except Exception:
            return 0

# Linha completa: DATA [DATA] [HORA] DOC HIST VALOR C/D SALDO [C/D]
FULL_LINE_RE = re.compile(r"""
^\s*
(?P<data>\d{2}/\d{2}(?:/\d{2}(?:\d{2})?)?)          # 03/01/2025 ou 03/01
(?:\s+\d{2}/\d{2}(?:/\d{2}(?:\d{2})?)?)?            # 03/01 (opcional)
(?:\s+\d{2}:\d{2})?                                 # 08:14 (opcional)
\s+(?P<doc>[\d./-]{6,})\s+                          # 000006727
(?P<hist>.*?)                                       # AZCX EL CD (lazy)
\s+(?P<valor>-?\s*(?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2})  # 273,24
\s*(?P<liber>[DC])\s+                               # C
(?P<saldo>-?\s*(?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2})     # 24.497,75
\s*(?P<sd>[DC])?                                    # D (opcional)
\s*$
""", re.IGNORECASE | re.VERBOSE)

# ---------- normalização final (6 colunas) ----------
_re_data   = re.compile(r"\b\d{2}/\d{2}(?:/\d{2}(?:\d{2})?)?\b")
_re_money  = re.compile(r"(?<!\d)-?\s*(?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2}")
_re_dc_any = re.compile(r"\b([DC])\b", re.I)
_re_doc    = re.compile(r"\b[\d./-]{6,}\b")
_SALDO_KEYS = ("SALDO", "SALDO ANTERIOR", "SALDO FINAL", "SALDO DO DIA", "SALDO ATUAL")


def drop_header_like_rows(df, log_entries=None, pdf_name=""):
    if df is None or df.empty:
        return df

    cleaned = []
    first_kept = False  # garante que a 1ª linha “de lançamento” fique

    for _, row in df.iterrows():
        raw = " ".join(map(str, row))
        S_UP = f" {raw.upper()} "
        S_COMPACT = re.sub(r"\s+", "", S_UP)

        # tem “cara” de lançamento?
        has_date  = bool(_re_data.search(raw))
        has_money = bool(_re_money.search(raw))
        looks_tx  = has_date or has_money or (" SALDO " in S_UP and has_money)

        # rodapé/identificação
        is_footer = (
            bool(re.search(r"P[ÁA]GINA\s+\d+\s+DE\s+\d+", S_UP)) or
            "CONSULTA REALIZADA" in S_UP or
            "TITULAR" in S_UP or "CPF/CNPJ" in S_UP or "PERÍODO" in S_UP or
            "PERIODO" in S_UP or "NOME DA UNIDADE" in S_UP
        )

        # cabeçalho típico de tabela
        header_tokens = ("DATA", "HIST", "HISTÓRICO", "VALOR", "SALDO", "DOC", "Nº DOC", "LIBER")
        token_hits = sum(1 for t in header_tokens if t in S_UP)
        is_header = token_hits >= 3 or any(sig in S_COMPACT for sig in [
            "SALDOVALORHISTÓRICONR.DOCDATAMOV",
            "SALDOVALORHISTORICONR.DOCDATAMOV",
            "EXTRATOHISTÓRICODACONTA", "EXTRATOHISTORICODACONTA",
            "CONTACORRENTEPESSOAJURIDICACAIXA"
        ])

        # regra de prioridade: se é a primeira linha “com cara de lançamento”, mantém SEMPRE
        if not first_kept and looks_tx:
            cleaned.append(row)
            first_kept = True
            if log_entries is not None:
                log_entries.append(f"PDF: {pdf_name}\nLinha Bruta: {raw.strip()}\nAção: Mantida (primeira válida da página)\n---\n")
            continue

        # descarta apenas se realmente for rodapé/cabeçalho E não parecer lançamento
        if (is_footer or is_header) and not looks_tx:
            if log_entries is not None:
                motivo = "Rodapé" if is_footer else "Cabeçalho"
                log_entries.append(f"PDF: {pdf_name}\nLinha Bruta: {raw.strip()}\nAção: Descartada\nMotivo: {motivo}\n---\n")
            continue

        cleaned.append(row)

    return pd.DataFrame(cleaned, columns=df.columns).reset_index(drop=True)

def _to_line(row) -> str:
    return " ".join([str(x).strip() for x in row if str(x).strip()])

def _pick_first(pattern, text):
    m = re.search(pattern, text)
    return m.group(0) if m else ""

def _remove_first(pattern, text):
    m = re.search(pattern, text)
    return (text[:m.start()] + text[m.end():]).strip() if m else text

def _find_dc_near(s: str, start: int, end: int) -> str:
    win_after = s[end:end+4]
    m = re.search(_re_dc_any, win_after)
    if m: return m.group(1).upper()
    win_before = s[max(0, start-2):start+1]
    m = re.search(_re_dc_any, win_before)
    return m.group(1).upper() if m else ""

def normalize_extrato(df_in: pd.DataFrame, log_entries=None, pdf_name="") -> pd.DataFrame:
    if df_in is None or df_in.empty:
        return pd.DataFrame(columns=COLUNAS_PADRAO)

    # junta colunas por linha num texto único
    linhas = df_in.iloc[:,0].astype(str).tolist() if df_in.shape[1] == 1 else df_in.apply(_to_line, axis=1).tolist()

    registros = []
    last_date = last_hist = last_doc = ""

    for linha in linhas:
        s = " ".join(linha.split())
        if not s:
            if log_entries is not None:
                log_entries.append(f"PDF: {pdf_name}\nLinha Bruta: {linha.strip()}\nAção: Descartada\nMotivo: Linha vazia.\n---\n")
            continue

        data = historico = documento = valor = liber = saldo = ""

        # 1) parser “linha inteira”
        m = FULL_LINE_RE.match(s)
        if m:
            data      = m.group("data") or ""
            documento = (m.group("doc") or "").strip()
            historico = " ".join((m.group("hist") or "").split())
            valor     = (m.group("valor") or "").replace("R$","").replace(" ","").strip()
            saldo     = (m.group("saldo") or "").replace("R$","").replace(" ","").strip()
            liber     = (m.group("liber") or "").upper()
            sd        = (m.group("sd") or "").upper()
            if sd:     # se o banco imprime D/C também no saldo, anexa
                saldo = f"{saldo} {sd}"
        else:
            # 2) caminho genérico
            S_UP = s.upper()
            data = _pick_first(_re_data, s)
            s2 = _remove_first(_re_data, s) if data else s

            matches = list(_re_money.finditer(s2))
            resto = s2
            if len(matches) >= 2:
                val_m, sal_m = matches[-2], matches[-1]
                valor = val_m.group().replace("R$","").replace(" ","").strip()
                saldo = sal_m.group().replace("R$","").replace(" ","").strip()
                liber = _find_dc_near(s2, val_m.start(), val_m.end())
                for a,b in sorted([(val_m.start(), val_m.end()), (sal_m.start(), sal_m.end())], reverse=True):
                    resto = (resto[:a] + resto[b:]).strip()
            elif len(matches) == 1:
                m1 = matches[0]
                num = m1.group().replace("R$","").replace(" ","").strip()
                liber = _find_dc_near(s2, m1.start(), m1.end())
                if any(k in S_UP for k in _SALDO_KEYS):
                    saldo = num
                else:
                    valor = num
                resto = (s2[:m1.start()] + s2[m1.end():]).strip()
            else:
                resto = s2

            documento = _pick_first(_re_doc, resto) or _pick_first(_re_doc, s2)
            historico = resto
            for m in matches:
                historico = historico.replace(m.group(), "")
            if documento:
                historico = historico.replace(documento, "")
            historico = re.sub(r"\b[DC]\b", "", historico, flags=re.I)
            historico = " ".join(historico.split())

            # 3) carry-forward p/ casos em que só sobram números (quebra de página)
            if not data and not historico and (valor or saldo):
                data = last_date
                historico = last_hist
                if not documento:
                    documento = last_doc

        # atualiza contexto
        if data:      last_date = data
        if historico: last_hist = historico
        if documento: last_doc  = documento

        current = [data, historico, documento, valor, liber, saldo]
        registros.append(current)

        if log_entries is not None:
            status = "Incluída" if any(str(c).strip() for c in current) else "Descartada"
            log_entries.append(f"PDF: {pdf_name}\nLinha Bruta: {s}\nAção: {status}\n---\n")

    out = pd.DataFrame(registros, columns=COLUNAS_PADRAO)

    # limpa “vazias”
    mask_vazia = out.apply(lambda r: "".join(r.values.astype(str)).strip()=="", axis=1)
    out = out[~mask_vazia]

    # se houver datas, mantém linhas com DATA ou com VALOR/SALDO
    is_data = out["DATA"].str.match(_re_data)
    if is_data.any():
        keep = is_data | out[["VALOR","SALDO"]].apply(lambda s: s.astype(str).str.strip()!="").any(axis=1)
        out = out[keep]

    return out.reset_index(drop=True)

def stitch_value_only_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    def _has_text(r):
        return bool(str(r.get("DATA","")).strip() or str(r.get("HISTORICO","")).strip())
    def _empty_vs(r):
        return (not str(r.get("VALOR","")).strip()) and (not str(r.get("SALDO","")).strip())
    to_drop = []
    n = len(df)
    for i in range(n):
        r = df.loc[i]
        is_orphan = (not str(r.get("DATA","")).strip() and not str(r.get("HISTORICO","")).strip() and (str(r.get("VALOR","")).strip() or str(r.get("SALDO","")).strip()))
        if not is_orphan:
            continue
        target = None
        if i+1 < n and _has_text(df.loc[i+1]) and _empty_vs(df.loc[i+1]):
            target = i+1
        elif i-1 >= 0 and _has_text(df.loc[i-1]) and _empty_vs(df.loc[i-1]):
            target = i-1
        elif i+1 < n and _has_text(df.loc[i+1]):
            target = i+1
        elif i-1 >= 0 and _has_text(df.loc[i-1]):
            target = i-1
        if target is not None:
            for col in ["VALOR","SALDO","LIBER","DOCUMENTO","DATA","HISTORICO"]:
                if not str(df.at[target, col]).strip() and str(r[col]).strip():
                    df.at[target, col] = r[col]
            to_drop.append(i)
    if to_drop:
        df = df.drop(index=to_drop).reset_index(drop=True)
    return df

# ---------- extratores ----------
def _import_tabula():
    try:
        import tabula
        return tabula
    except Exception:
        return None
def _import_pdfplumber():
    try:
        import pdfplumber
        return pdfplumber
    except Exception:
        return None
def _import_pdf2image():
    try:
        from pdf2image import convert_from_path
        return convert_from_path
    except Exception:
        return None
def _import_pytesseract():
    try:
        import pytesseract
        if TESSERACT_EXE:
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
        return pytesseract
    except Exception:
        return None
def extrair_com_tabula(pdf_path: str, lattice=True):
    tabula = _import_tabula()
    if tabula is None:
        return []

    n_pages = get_page_count(pdf_path)
    resultados = []

    # Se não conseguir contar páginas, cai no all (como estava)
    if n_pages <= 0:
        try:
            dfs = tabula.read_pdf(
                pdf_path, pages="all", multiple_tables=True,
                stream=not lattice, lattice=lattice, guess=True, silent=True
            )
            for d in dfs or []:
                d = limpar_df(d)
                if d.shape[1] >= MIN_COLUNAS_ACEITAS:
                    resultados.append(d)
            return resultados
        except Exception:
            return []

    # Força a segmentação: uma chamada por página
    for p in range(1, n_pages + 1):
        try:
            dfs_p = tabula.read_pdf(
                pdf_path, pages=str(p), multiple_tables=True,
                stream=not lattice, lattice=lattice, guess=True, silent=True
            )
            for d in dfs_p or []:
                d = limpar_df(d)
                if d.shape[1] >= MIN_COLUNAS_ACEITAS:
                    resultados.append(d)
        except Exception:
            continue

    return resultados
def extrair_com_pdfplumber(pdf_path: str):
    pdfplumber = _import_pdfplumber()
    if pdfplumber is None: return []
    tables_all = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                settings_lines = dict(vertical_strategy="lines", horizontal_strategy="lines", snap_tolerance=3)
                tbls = page.extract_tables(table_settings=settings_lines) or []
                if not tbls:
                    settings_text = dict(vertical_strategy="text", horizontal_strategy="text", text_tolerance=2)
                    tbls = page.extract_tables(table_settings=settings_text) or []
                for t in tbls:
                    df = limpar_df(pd.DataFrame(t))
                    if df.shape[1] >= MIN_COLUNAS_ACEITAS:
                        tables_all.append(df)
        return tables_all
    except Exception: return []

def extrair_com_ocr(pdf_path: str):
    convert_from_path = _import_pdf2image()
    pytesseract = _import_pytesseract()
    if convert_from_path is None or pytesseract is None:
        return []
    kwargs = {}
    if POPPLER_BIN and os.path.isdir(POPPLER_BIN):
        kwargs["poppler_path"] = POPPLER_BIN
    resultados = []
    try:
        images = convert_from_path(pdf_path, dpi=300, **kwargs)
        for img in images:
            tsv = pytesseract.image_to_data(img, lang="por+eng", output_type=pytesseract.Output.DATAFRAME)
            tsv = tsv.dropna(subset=["text"]).reset_index(drop=True)
            if tsv.empty: continue
            lines_by_y = tsv.groupby("line_num")
            linhas = []
            for _, line_group in lines_by_y:
                line_group = line_group.sort_values("left").reset_index(drop=True)
                line_text = " ".join(line_group["text"])
                linhas.append([line_text])
            if not linhas: continue
            df = limpar_df(pd.DataFrame(linhas))
            if df.shape[1] >= MIN_COLUNAS_ACEITAS:
                resultados.append(df)
        return resultados
    except Exception: return []

# ---------- pipeline por PDF (para ser usado no pool) ----------

def push_header_row_if_looks_like_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Se o extrator tiver promovido a primeira linha a 'header', recupera essa linha:
    - Se nomes das colunas tiverem cara de dados (data, dinheiro, doc ou muitos dígitos),
      transformamos os nomes das colunas em uma linha de dados no topo.
    - Em qualquer caso, padronizamos as colunas para 0..N-1 (sem headers 'reais').
    """
    if df is None or df.empty:
        return df

    df = df.copy()
    col_str = list(map(str, df.columns))

    # Sinais de que o "header" na verdade é uma linha de dados
    has_date   = any(_re_data.search(c)  for c in col_str)
    has_money  = any(_re_money.search(c) for c in col_str)
    has_doc    = any(_re_doc.search(c)   for c in col_str)
    many_digits= sum(bool(re.search(r"\d", c)) for c in col_str) >= max(2, len(col_str)//2)

    header_is_data = has_date or has_money or has_doc or many_digits

    # Corpo padronizado com colunas numéricas
    body = df.reset_index(drop=True)
    body.columns = range(body.shape[1])

    if header_is_data:
        first_row = pd.DataFrame([col_str])
        first_row.columns = range(body.shape[1])
        return pd.concat([first_row, body], ignore_index=True)

    return body


def processar_um_pdf(pdf_path: str):
    log_entries = []
    pdf_name = os.path.basename(pdf_path)
    
    candidatos = []
    dfs = extrair_com_ocr(pdf_path)
    if dfs: candidatos.append(("ocr", dfs))
    dfs = extrair_com_pdfplumber(pdf_path)
    if dfs: candidatos.append(("pdfplumber", dfs))
    dfs = extrair_com_tabula(pdf_path, lattice=True)
    if dfs: candidatos.append(("tabula_lattice", dfs))
    dfs = extrair_com_tabula(pdf_path, lattice=False)
    if dfs: candidatos.append(("tabula_stream", dfs))

    if not candidatos:
        log_entries.append(f"PDF: {pdf_name}\nLinha Bruta: N/A\nAção: Nenhum dado extraído\nMotivo: Nenhuma tabela detectada por nenhum método de extração.\n---\n")
        return (pdf_path, pd.DataFrame(columns=COLUNAS_PADRAO), log_entries)

    def score(dfs_):
        d = [limpar_df(x) for x in dfs_ if x is not None and not x.empty]
        if not d: return 0
        linhas = sum(len(x) for x in d)
        cols = max(x.shape[1] for x in d)
        return linhas * max(cols, 1)

    candidatos.sort(key=lambda item: score(item[1]), reverse=True)
    method_escolhido, dfs_escolhidos = candidatos[0]
    log_entries.append(f"PDF: {pdf_name}\nMétodo de Extração: {method_escolhido}\nAção: Selecionado\nMotivo: Foi o método com maior pontuação de linhas e colunas.\n---\n")

    # Em vez de colar tudo e normalizar de uma vez, tratamos por página/tabela
    paginas_limpa = []
    for df_pg in dfs_escolhidos:
        df_pg = limpar_df(df_pg)
        if df_pg is None or df_pg.empty:
            continue
        # garante que a 1ª linha nunca fica “presa” como header
        df_pg = push_header_row_if_looks_like_data(df_pg)

        # agora sim, filtra cabeçalho/rodapé sem risco de comer a 1ª linha válida
        df_pg = drop_header_like_rows(df_pg, log_entries, pdf_name)

        if not df_pg.empty:
            paginas_limpa.append(df_pg)

    # Normaliza CADA página separadamente (reinicia o estado entre páginas)
    normalizadas = []
    for df_pg in paginas_limpa:
        df_norm = normalize_extrato(df_pg, log_entries, pdf_name)
        df_norm = stitch_value_only_rows(df_norm)
        if not df_norm.empty:
            normalizadas.append(df_norm)

    if not normalizadas:
        log_entries.append(
            f"PDF: {pdf_name}\nLinha Bruta: N/A\nAção: Nenhum dado normalizado\nMotivo: Tabelas vazias após limpeza por página.\n---\n"
        )
        return (pdf_path, pd.DataFrame(columns=COLUNAS_PADRAO), log_entries)

    # Agora sim, concatena o resultado final já por página
    df = pd.concat(normalizadas, ignore_index=True, sort=False)
    return (pdf_path, df, log_entries)


# ---------- varredura e gravação com paralelismo ----------
def main():
    base = Path(PASTA_PDFS)
    if not base.exists():
        print(f"Erro: pasta não encontrada: {base}")
        sys.exit(1)
    pdfs = sorted(p for p in base.glob("*.pdf"))
    if not pdfs:
        print(f"Nenhum PDF em: {base}")
        sys.exit(0)

    num_processos = cpu_count()
    print(f"Processando...")

    with Pool(processes=num_processos) as pool:
        resultados = list(tqdm(pool.imap_unordered(processar_um_pdf, pdfs), total=len(pdfs)))

    destino_excel = base / NOME_EXCEL
    destino_log = base / NOME_LOG
    
    all_log_entries = []

    with pd.ExcelWriter(destino_excel, engine="openpyxl") as writer:
        for pdf, df, log_entries in resultados:
            print(f"✅ Processado: {os.path.basename(pdf)}")
            aba = pdf.stem[:31] or "PDF"
            if df.empty:
                pd.DataFrame({"AVISO":[f"Nenhuma tabela detectada em {pdf.name}"]}).to_excel(writer, sheet_name=aba, index=False)
            else:
                df.to_excel(writer, sheet_name=aba, index=False)
                try:
                    ws = writer.book[aba]
                    ws.freeze_panes = "A2"
                    for col in ws.columns:
                        maxlen = 10
                        for cell in col:
                            val = "" if cell.value is None else str(cell.value)
                            if len(val) > maxlen: maxlen = len(val)
                        ws.column_dimensions[col[0].column_letter].width = min(maxlen+2, 60)
                except Exception:
                    pass
            all_log_entries.extend(log_entries)

    with open(destino_log, "w", encoding="utf-8") as f:
        f.writelines(all_log_entries)

    print(f"✅ Concluído! Excel salvo em: {destino_excel}")
    print(f"📝 Log detalhado salvo em: {destino_log}")

if __name__ == "__main__":
    from multiprocessing import freeze_support
    freeze_support()
    main()