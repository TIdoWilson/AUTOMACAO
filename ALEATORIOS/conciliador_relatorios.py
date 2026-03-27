import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd


# =========================
# UI: seleção de arquivos
# =========================
def selecionar_pdfs():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    messagebox.showinfo(
        "Conciliação de PDFs",
        "Selecione o PDF do LIVRO (Registro de Entradas)."
    )
    livro = filedialog.askopenfilename(
        title="Selecione o LIVRO (Registro de Entradas)",
        filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
    )
    if not livro:
        raise SystemExit("Cancelado (Livro não selecionado).")

    messagebox.showinfo(
        "Conciliação de PDFs",
        "Selecione o PDF do RELATÓRIO TIPO 50."
    )
    tipo50 = filedialog.askopenfilename(
        title="Selecione o RELATÓRIO TIPO 50",
        filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
    )
    if not tipo50:
        raise SystemExit("Cancelado (Tipo 50 não selecionado).")

    return Path(livro), Path(tipo50)


# =========================
# Normalização
# =========================
def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()

def to_decimal_br(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == "" or s.upper() == "EX":
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        return float(s)
    except:
        return None

def norm_int(s):
    s = "" if s is None else str(s)
    s = re.sub(r"\D", "", s)
    return s if s else None

def norm_date_ddmmyyyy(s):
    s = "" if s is None else str(s).strip()
    m = re.search(r"(\d{2}/\d{2}/\d{4})", s)
    return m.group(1) if m else None

def fix_truncated_year(date_str: str, default_year="2026"):
    """
    Corrige datas como 05/01/202 (ano com 3 dígitos) -> 05/01/2026 por padrão.
    Ajuste default_year se seus PDFs forem de outro ano.
    """
    if date_str is None:
        return None
    date_str = date_str.strip()
    if re.fullmatch(r"\d{2}/\d{2}/\d{3}", date_str):
        return date_str + default_year[-1]
    return date_str


# =========================
# Leitura de tabelas (Camelot)
# =========================
def read_tables_camelot(pdf_path: Path):
    try:
        import camelot
        tables = camelot.read_pdf(str(pdf_path), pages="all", flavor="stream")
        dfs = []
        for t in tables:
            df = t.df.copy()
            df = df.dropna(how="all")
            if df.shape[0] >= 2 and df.shape[1] >= 6:
                dfs.append(df)
        return dfs
    except Exception:
        return []

def identificar_pdf(pdf_path: Path) -> str:
    """
    Tenta identificar se é LIVRO ou TIPO50 por palavras/padrões.
    """
    import pdfplumber
    with pdfplumber.open(str(pdf_path)) as pdf:
        txt = (pdf.pages[0].extract_text() or "")
    t = strip_accents(txt).upper()

    if "RELATORIO TIPO 50" in t or "TIPO 50" in t:
        return "TIPO50"
    if "LIVRO REGISTRO" in t or "REGISTRO DE ENTRADAS" in t:
        return "LIVRO"

    # fallback por padrão de linha:
    # Tipo50 costuma ter " T " ou " P " após número, CNPJ, UF, CFOP com ponto
    if re.search(r"\d{2}/\d{2}/\d{3,4}\s+\d+\s+[TP]\s+\d{11,14}", txt):
        return "TIPO50"

    # Livro costuma ter "NFe" / "CTe" e colunas "CFOP" "Emitente"
    if "NFE" in t or "CTE" in t or "EMITENTE" in t:
        return "LIVRO"

    return "DESCONHECIDO"


def parse_livro(pdf_path: Path) -> pd.DataFrame:
    """
    Tenta extrair o LIVRO.
    1) pdfplumber (texto)
    2) se não houver texto, OCR (pytesseract + pdf2image)
    """
    rows = []

    # ---------- Parser em "chunks" ----------
    def parse_text_to_rows(txt: str):
        # normaliza
        txt = (txt or "").replace("\n", " ")
        txt = clean_spaces(txt)

        # início de lançamento
        re_inicio = re.compile(r"(?P<data>\d{2}/\d{2}/\d{4})\s+(?P<esp>NFe|NF|CTe)\b", re.IGNORECASE)
        re_cfop = re.compile(r"\b(?P<cfop>(\d{4}|\d\.\d{3}))\b")
        re_valor = re.compile(r"\b(?P<valor>\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\b")

        starts = list(re_inicio.finditer(txt))
        if not starts:
            return

        for i, st in enumerate(starts):
            start_pos = st.start()
            end_pos = starts[i + 1].start() if i + 1 < len(starts) else len(txt)
            chunk = txt[start_pos:end_pos]

            data = st.group("data")
            esp = st.group("esp").upper()

            tail = chunk[st.end():]

            ints = re.findall(r"\b\d+\b", tail)
            if len(ints) < 2:
                continue

            serie = norm_int(ints[0])
            numero = norm_int(ints[1])

            mcfop = re_cfop.search(chunk)
            if not mcfop:
                continue
            cfop = norm_cfop(mcfop.group("cfop"))

            # pega o MAIOR valor com vírgula no chunk (normalmente o total)
            vals = [to_decimal_br(v) for v in re_valor.findall(chunk)]
            vals = [v for v in vals if v is not None]
            if not vals:
                continue
            valor = max(vals)

            rows.append({
                "data": data,
                "especie": esp,
                "serie": serie,
                "numero": numero,
                "cfop": cfop,
                "valor_livro": valor,
            })

    # ---------- 1) Tentativa: pdfplumber ----------
    import pdfplumber
    extracted_any_text = False
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t and t.strip():
                extracted_any_text = True
                parse_text_to_rows(t)

    # ---------- 2) Fallback OCR ----------
    if not rows and not extracted_any_text:
        try:
            import pytesseract
            from pdf2image import convert_from_path

            # converte páginas para imagens
            images = convert_from_path(str(pdf_path), dpi=300)

            for img in images:
                ocr_txt = pytesseract.image_to_string(img, lang="por")
                parse_text_to_rows(ocr_txt)

        except Exception:
            # se OCR não estiver configurado, vai retornar vazio e o main lida
            pass

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df = (df.groupby(["data", "serie", "numero", "cfop"], as_index=False)
            .agg(valor_livro=("valor_livro", "sum")))
    return df

def parse_tipo50(pdf_path: Path, default_year="2026") -> pd.DataFrame:
    """
    Extrai Tipo50:
    data, numero (nota), cfop, valor_tipo50
    (série normalmente não vem; se vier, dá pra adaptar depois)
    """
    rows = []
    dfs = read_tables_camelot(pdf_path)

    if dfs:
        for df in dfs:
            for i in range(len(df)):
                line = clean_spaces(" ".join(df.iloc[i].astype(str).tolist()))
                m = re.search(
                    r"(?P<data>\d{2}/\d{2}/\d{3,4})\s+"
                    r"(?P<num>\d+)\s+"
                    r"(?P<tip>[TP])\s+"
                    r"(?P<cnpj>[A-Z0-9\./-]+)\s+"
                    r"(?P<uf>[A-Z]{2})\s+"
                    r"(?P<cfop>\d\.\d{3})\s+"
                    r"(?P<aliq>[\d\.,]+)\s+"
                    r"(?P<total>[\d\.,]+)",
                    line
                )
                if m:
                    data = fix_truncated_year(m.group("data"), default_year=default_year)
                    rows.append({
                        "data": norm_date_ddmmyyyy(data),
                        "numero": norm_int(m.group("num")),
                        "cfop": norm_cfop(m.group("cfop")),  # 1.653 -> 1653
                        "valor_tipo50": to_decimal_br(m.group("total")),
                    })

    if not rows:
        import pdfplumber
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                txt = clean_spaces((page.extract_text() or "").replace("\n", " "))
                for m in re.finditer(
                    r"(\d{2}/\d{2}/\d{3,4})\s+(\d+)\s+([TP])\s+([A-Z0-9\./-]+)\s+([A-Z]{2})\s+(\d\.\d{3})\s+([\d\.,]+)\s+([\d\.,]+)",
                    txt
                ):
                    data = fix_truncated_year(m.group(1), default_year=default_year)
                    rows.append({
                        "data": norm_date_ddmmyyyy(data),
                        "numero": norm_int(m.group(2)),
                        "cfop": norm_cfop(m.group(6)),
                        "valor_tipo50": to_decimal_br(m.group(8)),
                    })

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df = (df.groupby(["data", "numero", "cfop"], as_index=False)
            .agg(valor_tipo50=("valor_tipo50", "sum")))
    return df


def conciliar_flex(df_livro: pd.DataFrame, df_tipo50: pd.DataFrame) -> pd.DataFrame:
    """
    Concilia em 2 passadas:
    1) data|serie|numero|cfop
    2) (restante) data|numero|cfop
    Sem tolerância: bateu só se valores iguais.
    """
    l = df_livro.copy()
    t = df_tipo50.copy()

    # garante tipos/string
    for c in ["data", "serie", "numero", "cfop"]:
        if c not in l.columns: l[c] = ""
    for c in ["data", "numero", "cfop"]:
        if c not in t.columns: t[c] = ""

    l["serie"] = l["serie"].fillna("").astype(str)
    t["serie"] = ""  # Tipo50 geralmente não tem série

    l["k1"] = l["data"].astype(str) + "|" + l["serie"] + "|" + l["numero"].astype(str) + "|" + l["cfop"].astype(str)
    t["k1"] = ""  # sem série, não dá para casar na 1a passada

    l["k2"] = l["data"].astype(str) + "|" + l["numero"].astype(str) + "|" + l["cfop"].astype(str)
    t["k2"] = t["data"].astype(str) + "|" + t["numero"].astype(str) + "|" + t["cfop"].astype(str)

    # merge pela chave k2 (que é a que realmente existe nos dois)
    m = l.merge(
        t[["k2", "valor_tipo50"]],
        left_on="k2",
        right_on="k2",
        how="outer"
    )

    m["valor_livro"] = m["valor_livro"].fillna(0.0)
    m["valor_tipo50"] = m["valor_tipo50"].fillna(0.0)
    m["diferenca"] = m["valor_livro"] - m["valor_tipo50"]

    # reconstroi colunas visíveis
    m["data"] = m["data"].fillna("")
    m["numero"] = m["numero"].fillna("")
    m["cfop"] = m["cfop"].fillna("")

    return m[["data", "serie", "numero", "cfop", "valor_livro", "valor_tipo50", "diferenca", "k2"]].sort_values(
        ["data", "numero", "cfop"], na_position="last"
    )


def print_conciliacao_exata(df_merge: pd.DataFrame):
    bateu = df_merge[(df_merge["valor_livro"] != 0) & (df_merge["valor_tipo50"] != 0) & (df_merge["diferenca"] == 0)]
    diverg = df_merge[(df_merge["valor_livro"] != 0) & (df_merge["valor_tipo50"] != 0) & (df_merge["diferenca"] != 0)]
    so_livro = df_merge[(df_merge["valor_livro"] != 0) & (df_merge["valor_tipo50"] == 0)]
    so_tipo = df_merge[(df_merge["valor_livro"] == 0) & (df_merge["valor_tipo50"] != 0)]

    print("\n==== RESUMO ====")
    print(f"Total linhas: {len(df_merge)}")
    print(f"Certo (verde) [dif=0]: {len(bateu)}")
    print(f"Divergente [dif!=0]: {len(diverg)}")
    print(f"Só no Livro: {len(so_livro)}")
    print(f"Só no Tipo 50: {len(so_tipo)}")

    def show(title, df):
        if df.empty:
            return
        print(f"\n==== {title} ({len(df)}) ====")
        with pd.option_context("display.max_rows", 5000, "display.width", 200):
            print(df[["data", "serie", "numero", "cfop", "valor_livro", "valor_tipo50", "diferenca"]].to_string(index=False))

    show("DIVERGÊNCIAS (dif != 0)", diverg)
    show("SÓ NO LIVRO", so_livro)
    show("SÓ NO TIPO 50", so_tipo)

def norm_cfop(v):
    """
    Converte CFOP como '1.653' ou '1653' ou ' 1653 ' -> '1653'
    """
    if v is None:
        return None
    s = str(v).strip()
    s = re.sub(r"\D", "", s)   # remove ponto e qualquer não-dígito
    if not s:
        return None
    # CFOP normalmente tem 4 dígitos; se vier com 3 por algum motivo, mantenha mesmo assim
    return s

def make_key(df, usar_cfop=False):
    d = df.copy()
    for col in ["data", "numero", "cfop"]:
        if col not in d.columns:
            d[col] = ""

    d["data"] = d["data"].fillna("").astype(str).str.strip()
    d["numero"] = d["numero"].fillna("").astype(str).str.strip()
    d["cfop"] = d["cfop"].apply(norm_cfop)

    if usar_cfop:
        d["key"] = d["data"] + "|" + d["numero"] + "|" + d["cfop"].fillna("")
    else:
        d["key"] = d["data"] + "|" + d["numero"]
    return d


def conciliar_por_nota(df_livro, df_tipo50, usar_cfop=False):
    l = make_key(df_livro, usar_cfop=usar_cfop)
    t = make_key(df_tipo50, usar_cfop=usar_cfop)

    l_agg = (l.groupby(["key"], as_index=False)
               .agg(data=("data", "first"),
                    numero=("numero", "first"),
                    cfop=("cfop", "first"),
                    valor_livro=("valor_livro", "sum")))

    t_agg = (t.groupby(["key"], as_index=False)
               .agg(data=("data", "first"),
                    numero=("numero", "first"),
                    cfop=("cfop", "first"),
                    valor_tipo50=("valor_tipo50", "sum")))

    m = l_agg.merge(t_agg, on="key", how="outer", suffixes=("_livro", "_tipo50"))

    m["data"] = m["data_livro"].fillna(m["data_tipo50"])
    m["numero"] = m["numero_livro"].fillna(m["numero_tipo50"])
    m["cfop"] = m["cfop_livro"].fillna(m["cfop_tipo50"])

    m["valor_livro"] = m["valor_livro"].fillna(0.0)
    m["valor_tipo50"] = m["valor_tipo50"].fillna(0.0)
    m["diferenca"] = m["valor_livro"] - m["valor_tipo50"]

    return m[["data", "numero", "cfop", "valor_livro", "valor_tipo50", "diferenca", "key"]]\
            .sort_values(["data", "numero"], na_position="last")

def print_diferencas_terminal(df_merge, tolerancia=0.10):
    """
    Mostra apenas:
    - divergências com |dif| > tolerância (não-verde)
    - só no livro
    - só no tipo 50
    """
    m = df_merge.copy()

    so_livro = m[(m["valor_livro"] != 0) & (m["valor_tipo50"] == 0)]
    so_tipo  = m[(m["valor_livro"] == 0) & (m["valor_tipo50"] != 0)]
    diverg   = m[(m["valor_livro"] != 0) & (m["valor_tipo50"] != 0) & (m["diferenca"].abs() > tolerancia)]
    bateu    = m[(m["valor_livro"] != 0) & (m["valor_tipo50"] != 0) & (m["diferenca"].abs() <= tolerancia)]

    print("\n==== RESUMO ====")
    print(f"Total linhas: {len(m)}")
    print(f"Bateu (<= {tolerancia:.2f}): {len(bateu)}")
    print(f"Divergente (> {tolerancia:.2f}): {len(diverg)}")
    print(f"Só no Livro: {len(so_livro)}")
    print(f"Só no Tipo 50: {len(so_tipo)}")

    def show(title, df):
        if df.empty:
            return
        print(f"\n==== {title} ({len(df)}) ====")
        with pd.option_context("display.max_rows", 5000, "display.width", 200):
            print(df[["data", "numero", "cfop", "valor_livro", "valor_tipo50", "diferenca"]].to_string(index=False))

    show("DIVERGÊNCIAS (não-verde)", diverg)
    show("SÓ NO LIVRO", so_livro)
    show("SÓ NO TIPO 50", so_tipo)

def main():
    while True:
        a, b = selecionar_pdfs()
        tipo_a = identificar_pdf(a)
        tipo_b = identificar_pdf(b)

        if {tipo_a, tipo_b} != {"LIVRO", "TIPO50"}:
            print(f"\nAVISO: não consegui identificar corretamente os PDFs.")
            print(f"Arquivo A: {a.name} -> {tipo_a}")
            print(f"Arquivo B: {b.name} -> {tipo_b}")
            print("Selecione novamente.\n")
            continue

        livro_pdf = a if tipo_a == "LIVRO" else b
        tipo50_pdf = a if tipo_a == "TIPO50" else b

        print(f"Livro selecionado:  {livro_pdf}")
        print(f"Tipo 50 selecionado:{tipo50_pdf}")

        df_livro = parse_livro(livro_pdf)
        if df_livro.empty:
            print("\nAVISO: Não consegui extrair dados do LIVRO. Selecione novamente.\n")
            continue

        df_tipo50 = parse_tipo50(tipo50_pdf, default_year="2026")
        if df_tipo50.empty:
            print("\nAVISO: Não consegui extrair dados do Tipo 50. Selecione novamente.\n")
            continue

        df_merge = conciliar_flex(df_livro, df_tipo50)
        print_conciliacao_exata(df_merge)

        break

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nERRO: {e}")
        sys.exit(1)
