# excel_assistido.py
import time
import logging
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import pythoncom
import win32com.client as win32

# Log opcional (na mesma pasta do script)
LOGFILE = Path(__file__).with_name("_assistido.log")
logging.basicConfig(
    filename=str(LOGFILE),
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

SLEEP = 0.4  # velocidade do “passo a passo” ao montar a planilha


def _fmt_day_brazil(d):
    """Aceita date / datetime / pandas Timestamp e devolve 'dd/mm/aaaa' (texto)."""
    if isinstance(d, pd.Timestamp):
        d = d.to_pydatetime()
    if isinstance(d, datetime):
        d = d.date()
    if not isinstance(d, date):
        raise TypeError(f"Dia inválido: {type(d)}")
    return f"{d.day:02d}/{d.month:02d}/{d.year:04d}"


def open_excel_and_prepare(xlsx_path: Path):
    """Abre Excel e o arquivo; devolve (excel, workbook)."""
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")

    # Alguns ambientes bloqueiam setar essas propriedades → tentamos sem quebrar
    try:
        excel.Visible = True
    except Exception:
        pass
    try:
        excel.DisplayAlerts = False
    except Exception:
        pass

    wb = excel.Workbooks.Open(str(xlsx_path))
    time.sleep(SLEEP)
    return excel, wb


def open_excel_and_present_day(xlsx_path: Path, df_day: pd.DataFrame, dia, slow_ms: int = 800):
    """
    Cria uma nova aba com os lançamentos do 'dia' (já fornecidos em df_day),
    ordena por Histórico, calcula totais e deixa tudo visível para inspeção.

    Parâmetros esperados em df_day (qualquer uma das variações por coluna):
      Data:      ["Data", "data"]
      Histórico: ["Historico", "Histórico", "historico"]
      Débito:    ["Debito", "Dêbito", "Débito", "debito"]
      Crédito:   ["Credito", "Crédito", "credito"]
    """
    # normaliza dia em texto
    dia_txt = _fmt_day_brazil(pd.to_datetime(dia))
    dia_dt = pd.to_datetime(dia).date()
    logging.info(f"Assistido: preparando aba para o dia {dia_txt}")

    # helper para achar nomes de colunas equivalentes
    def _get_col(df, keys):
        for k in keys:
            if k in df.columns:
                return k
        return None

    c_data = _get_col(df_day, ["Data", "data"])
    c_hist = _get_col(df_day, ["Historico", "Histórico", "historico"])
    c_deb  = _get_col(df_day, ["Debito", "Dêbito", "Débito", "debito"])
    c_cred = _get_col(df_day, ["Credito", "Crédito", "credito"])

    if not all([c_data, c_hist, c_deb, c_cred]):
        raise ValueError("df_day precisa conter colunas Data/Historico/Debito/Credito.")

    # garante tipos numéricos
    def _to_float(x):
        try:
            if pd.isna(x):
                return 0.0
            return float(x)
        except Exception:
            try:
                s = str(x).strip().replace(".", "").replace(",", ".")
                return float(s)
            except Exception:
                return 0.0

    dfw = df_day[[c_data, c_hist, c_deb, c_cred]].copy()
    # ordena pelo histórico pra “juntar” notas
    dfw.sort_values(c_hist, inplace=True, kind="mergesort")

    # abre excel / arquivo
    excel, wb = open_excel_and_prepare(Path(xlsx_path))

    try:
        # cria nome de aba ÚNICO
        base_name = f"Assistido_{dia_dt.strftime('%d%m')}"
        existing = {s.Name for s in wb.Sheets}
        new_name = base_name
        idx = 2
        while new_name in existing:
            new_name = f"{base_name}_{idx}"
            idx += 1

        ws = wb.Sheets.Add()
        ws.Name = new_name
        time.sleep(SLEEP)

        # cabeçalhos
        headers = ["Data", "Histórico", "Débito", "Crédito"]
        for j, h in enumerate(headers, start=1):
            ws.Cells(1, j).Value = h
        time.sleep(SLEEP / 2)

        # escreve dados
        n = len(dfw)
        for i, row in enumerate(dfw.itertuples(index=False), start=2):
            # Data
            try:
                dt_txt = pd.to_datetime(getattr(row, c_data)).strftime("%d/%m/%Y")
            except Exception:
                dt_txt = str(getattr(row, c_data))
            ws.Cells(i, 1).Value = dt_txt

            # Histórico
            hist = getattr(row, c_hist)
            ws.Cells(i, 2).Value = "" if pd.isna(hist) else str(hist)

            # Débito / Crédito
            ws.Cells(i, 3).Value = _to_float(getattr(row, c_deb))
            ws.Cells(i, 4).Value = _to_float(getattr(row, c_cred))

        # formatos e totais
        if n > 0:
            ws.Range(ws.Cells(2, 3), ws.Cells(1 + n, 4)).NumberFormat = "#,##0.00"

        ws.Cells(1, 6).Value = "Σ Débito"
        ws.Cells(1, 7).Value = "Σ Crédito"
        ws.Cells(1, 8).Value = "Δ (D-C)"
        if n > 0:
            ws.Cells(2, 6).FormulaLocal = f"=SOMA(C2:C{1+n})"
            ws.Cells(2, 7).FormulaLocal = f"=SOMA(D2:D{1+n})"
            ws.Cells(2, 8).FormulaLocal = f"=F2-G2"

        # re-ordena visualmente (já veio ordenado, mas o Sort ajuda na UI)
        if n > 1:
            ws.Range(ws.Cells(2, 1), ws.Cells(1 + n, 4)).Sort(Key1=ws.Range("B2"), Order1=1, Header=0)

        # ajustes visuais
        ws.Columns("A:D").AutoFit()
        try:
            excel.ActiveWindow.SplitRow = 1
            excel.ActiveWindow.FreezePanes = True
        except Exception:
            pass

        ws.Cells(4, 6).Value = "Assistido: revise pares, triplas e quadras"
        ws.Cells(5, 6).Value = f"Dia: {dia_txt}"

        # status bar com totais
        try:
            deb_sum = ws.Application.WorksheetFunction.Sum(ws.Range(f"C2:C{1+n}")) if n else 0.0
            cred_sum = ws.Application.WorksheetFunction.Sum(ws.Range(f"D2:D{1+n}")) if n else 0.0
            ws.Application.StatusBar = f"[ASSISTIDO] {dia_txt} | Débito: {deb_sum:.2f} | Crédito: {cred_sum:.2f} | Dif: {(deb_sum-cred_sum):.2f}"
        except Exception:
            ws.Application.StatusBar = f"[ASSISTIDO] {dia_txt}"

        # passo a passo mais lento para visualização
        time.sleep(max(0, int(slow_ms)) / 1000.0)

        logging.info(f"Assistido: aba '{new_name}' criada para {dia_txt} (Excel deixado aberto).")

    except Exception as e:
        logging.exception(f"Assistido falhou para {dia_txt}: {e}")
        print(f"[ASSISTIDO][ERRO] {dia_txt}: {e}")
        # mantém Excel aberto para inspeção manual


# ------------ UTILITÁRIOS OPCIONAIS ---------------
def assist_days(xlsx_path: Path, days):
    """Cria uma aba para cada dia fornecido (cada aba com dados do próprio dia)."""
    for d in days:
        logging.info(f"Assistido: solicitado dia {d} — assist_days requer df_day por dia.")
        time.sleep(0.5)
