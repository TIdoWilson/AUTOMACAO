import argparse
import json
import time
from datetime import datetime, date
from pathlib import Path

import pyautogui as pag
import uiautomation as uia
import msvcrt

BASE_DIR = Path(__file__).resolve().parent
COORDS_PATH = Path(__file__).with_suffix(".json")

DATA_DIA_FIXO = 20
EXPORT_DIR = Path(
    r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\1 - Pro Labore"
)
EMPRESAS_SLK = EXPORT_DIR / "empresa.slk"
EMPRESAS_XLSX = EXPORT_DIR / "empresas.xlsx"

WAIT_SHORT = 0.2
WAIT_MED = 0.8
WAIT_LONG = 2.0
WAIT_1_HOUR = 5 * 60
WAIT_10_MIN = 90
WAIT_2_SEC = 1.0
WAIT_0_5_SEC = 0.5
WAIT_TINY = 0.05
WAIT_STEP = 0.1

pag.PAUSE = 0.1
pag.FAILSAFE = False

LAST_ACTION = None


def load_coords(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de coordenadas nao encontrado: {path}")
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    required = [
        "data_sistema",
        "mes_ano",
        "marcar_por_grupo",
        "agrupamento",
        "selecionar",
        "marcar_todas",
        "gerar_periodicos",
        "limpar",
        "campo_empresa",
        "aba_com_erro",
        "marcar_todas_erro",
        "reenviar",
        "confirmar_reenviar",
        "imprimir_inicial",
        "grupos",
        "grupos2",
        "ordem_numerica",
        "imprimir_final",
        "salvar",
        "fechar_impressao",
    ]
    for key in required:
        if key not in data:
            raise KeyError(f"Coordenada ausente no JSON: {key}")
        coord = data[key]
        if (
            not isinstance(coord, (list, tuple))
            or len(coord) != 2
            or not all(isinstance(v, (int, float)) for v in coord)
        ):
            raise ValueError(f"Coordenada invalida para {key}: {coord}")
    return data


def click_at(coord):
    x, y = coord
    pag.moveTo(x, y, duration=0.1)
    pag.click()
    time.sleep(WAIT_STEP)



def record_action(action: dict):
    global LAST_ACTION
    LAST_ACTION = action

def do_click(coord):
    click_at(coord)
    record_action({"type": "click", "coord": coord})

def do_type(coord, text):
    click_at(coord)
    time.sleep(WAIT_SHORT)
    pag.hotkey("ctrl", "a")
    pag.press("backspace")
    pag.typewrite(text, interval=0.02)
    record_action({"type": "type", "coord": coord, "text": text})
    time.sleep(WAIT_STEP)


def click_and_type(coord, text):
    click_at(coord)
    time.sleep(WAIT_SHORT)
    pag.typewrite(text, interval=0.02)
    time.sleep(WAIT_STEP)


def alt_sequence(*keys):
    pag.keyDown("alt")
    for k in keys:
        pag.press(k)
        time.sleep(0.05)
    pag.keyUp("alt")
    time.sleep(WAIT_STEP)

def error_screenshot():
    ts = datetime.now().strftime("%H%M%S - %d-%m-%Y")
    out_dir = Path(
        r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\1 - Pro Labore\Prints NMR erros"
    )
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"Erro {ts}.png"
    img = pag.screenshot()
    img.save(out)
    print(f"[erro] print salvo: {out}")

def maybe_handle_error_dialog():
    wnd = uia.WindowControl(Name="Aviso do Sistema")
    try:
        if not wnd.Exists(0.2):
            return False
    except Exception:
        return False

    try:
        wnd.SetFocus()
    except Exception:
        pass

    error_screenshot()

    btn_target = None
    try:
        for c in wnd.GetChildren():
            if c.ControlTypeName != "ButtonControl":
                continue
            name = getattr(c, "Name", "") or ""
            if name.strip().upper() == "OK":
                btn_target = c
                break
            if btn_target is None:
                btn_target = c
    except Exception:
        btn_target = None

    if btn_target:
        try:
            inv = btn_target.GetInvokePattern()
            if inv:
                inv.Invoke()
                return True
        except Exception:
            pass
        try:
            btn_target.Click()
            return True
        except Exception:
            pass

        try:
            rect = btn_target.BoundingRectangle
            cx = int((rect.left + rect.right) / 2)
            cy = int((rect.top + rect.bottom) / 2)
            pag.moveTo(cx, cy, duration=0.1)
            pag.click()
            return True
        except Exception:
            pass

    pag.press("enter")

    # Repete o ultimo comando para evitar perda ou duplicacao parcial.
    if LAST_ACTION:
        print("[info] repetir ultimo comando")
        if LAST_ACTION.get("type") == "click":
            click_at(LAST_ACTION["coord"])
        elif LAST_ACTION.get("type") == "type":
            click_at(LAST_ACTION["coord"])
            time.sleep(WAIT_SHORT)
            pag.hotkey("ctrl", "a")
            pag.press("backspace")
            pag.typewrite(LAST_ACTION["text"], interval=0.02)
    return True

def wait_or_skip(total_seconds, label):
    if total_seconds <= 0:
        return
    print(f"[info] esperar {label}")
    step = 0.2
    waited = 0.0
    while waited < total_seconds:
        if msvcrt.kbhit():
            ch = msvcrt.getwch()
            if ch == "\r":
                print("[ok] espera pulada")
                return
        maybe_handle_error_dialog()
        time.sleep(step)
        waited += step


def focus_folha_pagamento():
    wnd = uia.WindowControl(Name="Folha de Pagamento", ClassName="TfrmPrincipal")
    if not wnd.Exists(2):
        return False
    try:
        wnd.SetFocus()
        return True
    except Exception:
        return False


def confirm_data_sistema(coord):
    focus_folha_pagamento()
    time.sleep(WAIT_SHORT)
    click_at(coord)
    time.sleep(WAIT_SHORT)
    pag.press("enter")
    pag.press("enter")


def parse_slk_empresas(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo SLK nao encontrado: {path}")
    cells = {}
    max_row = 0
    max_col = 0
    with path.open("r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line.startswith("C;"):
                continue
            parts = line.split(";")[1:]
            row = col = None
            val = None
            for p in parts:
                if p.startswith("X"):
                    try:
                        col = int(p[1:])
                    except Exception:
                        col = None
                elif p.startswith("Y"):
                    try:
                        row = int(p[1:])
                    except Exception:
                        row = None
                elif p.startswith("K"):
                    val = p[1:]
            if row is None or col is None or val is None:
                continue
            if val.startswith('"') and val.endswith('"'):
                val = val[1:-1]
            cells[(row, col)] = val
            if row > max_row:
                max_row = row
            if col > max_col:
                max_col = col

    def normalize_company_value(v):
        s = str(v).strip()
        if not s:
            return None
        if s.replace(".", "").isdigit():
            if "." in s:
                left, right = s.split(".", 1)
                if set(right) <= {"0"}:
                    s = left
            return s
        return None

    vals = []
    stop = False
    for r in range(1, max_row + 1):
        if r <= 2:
            continue
        row_vals = [cells[(r, c)] for c in range(1, max_col + 1) if (r, c) in cells]
        for rv in row_vals:
            if "rotina" in str(rv).lower():
                stop = True
                break
        for c in range(1, max_col + 1):
            if 2 <= c <= 7:
                continue
            if (r, c) not in cells:
                continue
            nv = normalize_company_value(cells[(r, c)])
            if nv:
                vals.append(nv)
        if stop:
            break
    return vals


def write_xlsx_single_col(path: Path, values: list[str]) -> None:
    import zipfile
    from xml.sax.saxutils import escape

    rows_xml = []
    for i, val in enumerate(values, 1):
        cell = (
            f'<row r="{i}"><c r="A{i}" t="inlineStr"><is><t>'
            f"{escape(val)}</t></is></c></row>"
        )
        rows_xml.append(cell)

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>"
        + "".join(rows_xml)
        + "</sheetData></worksheet>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Planilha1" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def convert_slk_to_empresas_xlsx():
    nums = parse_slk_empresas(EMPRESAS_SLK)
    if not nums:
        raise RuntimeError("Nenhum numero encontrado no SLK.")
    write_xlsx_single_col(EMPRESAS_XLSX, nums)
    print(f"[ok] empresas.xlsx atualizado: {len(nums)}")


def run_post_periodicos(coords, empresas):
    # LIMPAR inicial apos gerar periodicos
    maybe_handle_error_dialog()
    do_click(coords["limpar"])
    time.sleep(0.1)

    # Fluxo por empresa
    for empresa in empresas:
        maybe_handle_error_dialog()
        do_click(coords["marcar_por_grupo"])

        do_type(coords["campo_empresa"], empresa)

        do_click(coords["selecionar"])
        time.sleep(1.0)

        do_click(coords["aba_com_erro"])
        time.sleep(WAIT_TINY)

        do_click(coords["marcar_todas_erro"])
        time.sleep(WAIT_TINY)
        do_click(coords["reenviar"])

        time.sleep(WAIT_2_SEC)
        do_click(coords["confirmar_reenviar"])
        time.sleep(WAIT_0_5_SEC)

        # LIMPAR antes da proxima empresa
        maybe_handle_error_dialog()
        do_click(coords["limpar"])
        time.sleep(0.1)


def load_empresas(path: Path, sheet: str | None, col: str, skip_header: bool) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de empresas nao encontrado: {path}")

    if path.suffix.lower() in [".csv", ".txt"]:
        import csv
        with path.open("r", encoding="utf-8") as f:
            reader = csv.reader(f)
            vals = []
            for row in reader:
                if not row:
                    continue
                val = str(row[0]).strip()
                if val:
                    vals.append(val)
    else:
        try:
            import openpyxl
            from openpyxl.utils import column_index_from_string
        except Exception as e:
            # Fallback sem openpyxl: ler XML do XLSX (zip)
            import zipfile
            import xml.etree.ElementTree as ET
            import re

            def col_to_idx(c):
                c = c.strip().upper()
                if not c or not re.match(r"^[A-Z]+$", c):
                    raise ValueError(f"Coluna invalida: {c}")
                idx = 0
                for ch in c:
                    idx = idx * 26 + (ord(ch) - ord("A") + 1)
                return idx

            def idx_to_col(idx):
                s = ""
                while idx > 0:
                    idx, r = divmod(idx - 1, 26)
                    s = chr(r + ord("A")) + s
                return s

            def pick_sheet_name(zf):
                wb_xml = ET.fromstring(zf.read("xl/workbook.xml"))
                ns = {"a": wb_xml.tag.split("}")[0].strip("{")}
                sheets = wb_xml.findall(".//a:sheets/a:sheet", ns)
                sheet_map = {}
                rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
                rns = {"r": rels.tag.split("}")[0].strip("{")}
                rel_map = {r.get("Id"): r.get("Target") for r in rels.findall("r:Relationship", rns)}
                for sh in sheets:
                    name = sh.get("name")
                    rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    target = rel_map.get(rid)
                    if target and not target.startswith("xl/"):
                        target = f"xl/{target}"
                    sheet_map[name] = target
                if sheet and sheet in sheet_map:
                    return sheet_map[sheet]
                # fallback: primeiro sheet
                if sheets:
                    first = sheets[0]
                    rid = first.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    target = rel_map.get(rid)
                    if target and not target.startswith("xl/"):
                        target = f"xl/{target}"
                    return target
                return None

            with zipfile.ZipFile(path, "r") as zf:
                shared = {}
                if "xl/sharedStrings.xml" in zf.namelist():
                    ss_xml = ET.fromstring(zf.read("xl/sharedStrings.xml"))
                    ss_ns = {"a": ss_xml.tag.split("}")[0].strip("{")}
                    for i, si in enumerate(ss_xml.findall("a:si", ss_ns)):
                        text = "".join(t.text or "" for t in si.findall(".//a:t", ss_ns))
                        shared[i] = text

                sheet_xml = pick_sheet_name(zf)
                if not sheet_xml:
                    raise RuntimeError("Nenhuma planilha encontrada no XLSX.")
                root = ET.fromstring(zf.read(sheet_xml))
                ns = {"a": root.tag.split("}")[0].strip("{")}

                target_col = idx_to_col(col_to_idx(col))
                vals = []
                for row in root.findall(".//a:sheetData/a:row", ns):
                    for c in row.findall("a:c", ns):
                        ref = c.get("r", "")
                        if not ref.startswith(target_col):
                            continue
                        val = ""
                        if c.get("t") == "inlineStr":
                            tnode = c.find(".//a:t", ns)
                            if tnode is not None and tnode.text:
                                val = tnode.text
                        else:
                            v = c.find("a:v", ns)
                            if v is None:
                                continue
                            val = v.text or ""
                            if c.get("t") == "s":
                                val = shared.get(int(val), "")
                        s = str(val).strip()
                        if s:
                            vals.append(s)
                        break
        else:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb[sheet] if sheet else wb.active
            col_idx = column_index_from_string(col)
            vals = []
            for row in ws.iter_rows(min_row=1, min_col=col_idx, max_col=col_idx, values_only=True):
                v = row[0]
                if v is None:
                    continue
                s = str(v).strip()
                if s:
                    vals.append(s)

    if skip_header and vals:
        vals = vals[1:]
    return vals


def screenshot_here():
    ts = datetime.now().strftime("%H%M%S - %d-%m-%Y")
    out_dir = Path(
        r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\1 - Pro Labore\Prints NMR erros"
    )
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"Print {ts}.png"
    img = pag.screenshot()
    img.save(out)
    print(f"[ok] print salvo: {out}")


def main(
    skip_gerar_periodicos=False,
    empresas_file=None,
    empresas_sheet=None,
    empresas_col="A",
    empresas_skip_header=False,
    preflow_once_per_month=False,
    preflow_only=False,
    post_only=False,
):
    coords = load_coords(COORDS_PATH)
    empresas_path = Path(empresas_file) if empresas_file else (BASE_DIR / "empresas.xlsx")

    if post_only:
        empresas = load_empresas(empresas_path, empresas_sheet, empresas_col, empresas_skip_header)
        run_post_periodicos(coords, empresas)
        return

    print("[info] foco em 3s")
    time.sleep(3)

    focus_folha_pagamento()
    time.sleep(WAIT_SHORT)

    # DATA DO SISTEMA: 15/{mes atual}/{ano atual}
    hoje = datetime.now()
    data_sistema = date(hoje.year, hoje.month, DATA_DIA_FIXO).strftime("%d/%m/%Y")
    click_at(coords["data_sistema"])
    time.sleep(WAIT_SHORT)
    pag.press("backspace", presses=10, interval=0.02)
    pag.typewrite(data_sistema, interval=0.02)
    pag.press("enter")
    pag.press("enter")
    record_action({"type": "type", "coord": coords["data_sistema"], "text": data_sistema})
    time.sleep(WAIT_MED)

    # Fluxo extra: impressao/geracao da lista de empresas
    run_preflow = True
    if preflow_once_per_month:
        stamp = BASE_DIR / ".preflow_last_run.txt"
        ym = f"{hoje.year:04d}-{hoje.month:02d}"
        if stamp.exists() and stamp.read_text(encoding="utf-8").strip() == ym:
            run_preflow = False
        else:
            stamp.write_text(ym, encoding="utf-8")

    opened_main = False
    if run_preflow:
        maybe_handle_error_dialog()
        alt_sequence("d", "u")
        time.sleep(WAIT_SHORT)

        do_click(coords["imprimir_inicial"])
        time.sleep(WAIT_SHORT)

        do_type(coords["grupos"], "1")
        time.sleep(WAIT_SHORT)

        do_click(coords["grupos2"])
        time.sleep(WAIT_SHORT)

        do_click(coords["ordem_numerica"])
        time.sleep(WAIT_SHORT)

        do_click(coords["imprimir_final"])
        time.sleep(4)

        do_click(coords["salvar"])
        time.sleep(1)

        pag.hotkey("ctrl", "l")
        time.sleep(WAIT_SHORT)
        pag.typewrite(str(EXPORT_DIR), interval=0.01)
        pag.press("enter")
        time.sleep(WAIT_SHORT)

        pag.hotkey("alt", "t")
        time.sleep(WAIT_SHORT)
        pag.typewrite("slk", interval=0.02)
        pag.press("enter")
        time.sleep(WAIT_SHORT)

        pag.hotkey("alt", "n")
        time.sleep(WAIT_SHORT)
        pag.typewrite("empresa.slk", interval=0.02)
        pag.press("enter")
        time.sleep(WAIT_MED)

        convert_slk_to_empresas_xlsx()

        do_click(coords["fechar_impressao"])
        time.sleep(1)
        focus_folha_pagamento()
        alt_sequence("e", "m", "g")
        opened_main = True

    if preflow_only:
        print("[ok] preflow concluido")
        return

    empresas = load_empresas(empresas_path, empresas_sheet, empresas_col, empresas_skip_header)

    # ALT + E + M + G (segurar ALT e apertar os outros na ordem)
    if not opened_main:
        alt_sequence("e", "m", "g")

    # MES/ANO: {mes atual/ano atual}
    mes_ano = hoje.strftime("%m/%Y")
    click_at(coords["mes_ano"])
    time.sleep(WAIT_SHORT)
    pag.press("backspace", presses=7, interval=0.02)
    pag.typewrite(mes_ano, interval=0.02)
    record_action({"type": "type", "coord": coords["mes_ano"], "text": mes_ano})
    time.sleep(WAIT_MED)

    # MARCAR "POR GRUPO"
    do_click(coords["marcar_por_grupo"])
    time.sleep(WAIT_SHORT)

    # AGRUPAMENTO: 1
    do_type(coords["agrupamento"], "1")
    time.sleep(WAIT_SHORT)

    # SELECIONAR
    do_click(coords["selecionar"])
    wait_or_skip(WAIT_10_MIN, "3 minutos")

    # SELECIONAR "MARCAR TODAS"
    do_click(coords["marcar_todas"])
    time.sleep(WAIT_MED)

    # GERAR PERIODICOS
    if skip_gerar_periodicos:
        print("[info] pular gerar periodicos")
    else:
        do_click(coords["gerar_periodicos"])
        wait_or_skip(WAIT_1_HOUR, "15 minutos")
        screenshot_here()

    run_post_periodicos(coords, empresas)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--skip-gerar-periodicos",
        action="store_true",
        help="Executa o fluxo sem clicar em Gerar Periodicos (modo teste).",
    )
    ap.add_argument(
        "--skip-emitir",
        action="store_true",
        help="Alias de --skip-gerar-periodicos.",
    )
    ap.add_argument(
        "--empresas-file",
        default=None,
        help="Arquivo com lista de empresas (.xlsx/.csv). Padrao: empresas.xlsx na mesma pasta.",
    )
    ap.add_argument(
        "--empresas-sheet",
        default=None,
        help="Nome da aba do Excel (padrao: aba ativa).",
    )
    ap.add_argument(
        "--empresas-col",
        default="A",
        help="Coluna com os numeros das empresas (padrao: A).",
    )
    ap.add_argument(
        "--empresas-skip-header",
        action="store_true",
        help="Ignora a primeira linha do arquivo.",
    )
    ap.add_argument(
        "--preflow-once-per-month",
        action="store_true",
        help="Executa o fluxo inicial apenas uma vez por mes.",
    )
    ap.add_argument(
        "--preflow-only",
        action="store_true",
        help="Executa apenas o fluxo inicial de impressao/lista de empresas.",
    )
    ap.add_argument(
        "--post-only",
        action="store_true",
        help="Executa apenas o pos-gerar-periodicos (limpar + loop de empresas).",
    )
    args = ap.parse_args()
    skip_gerar = args.skip_gerar_periodicos or args.skip_emitir
    main(
        skip_gerar_periodicos=skip_gerar,
        empresas_file=args.empresas_file,
        empresas_sheet=args.empresas_sheet,
        empresas_col=args.empresas_col,
        empresas_skip_header=args.empresas_skip_header,
        preflow_once_per_month=args.preflow_once_per_month,
        preflow_only=args.preflow_only,
        post_only=args.post_only,
    )
