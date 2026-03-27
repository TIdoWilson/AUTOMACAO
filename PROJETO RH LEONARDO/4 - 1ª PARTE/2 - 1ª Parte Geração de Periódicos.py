import argparse
import json
import re
import time
from datetime import date, datetime
from pathlib import Path

import msvcrt
import pyautogui as pag
import uiautomation as uia

COORDS_PATH = Path(__file__).with_suffix(".json")

DATA_DIA_FIXO = 20
DEFAULT_AGRUPAMENTO = "6;13;14"
OUTPUT_ROOT = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\1ª PARTE")
EXCEL_DIR = Path(__file__).resolve().parent / "Grupos por mês Excel"

WAIT_SHORT = 0.2
WAIT_MED = 0.8
WAIT_1_HOUR = 60 * 60
WAIT_SELECT = 90
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
    with path.open("r", encoding="utf-8-sig") as f:
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


def alt_sequence(*keys):
    pag.keyDown("alt")
    for k in keys:
        pag.press(k)
        time.sleep(0.05)
    pag.keyUp("alt")
    time.sleep(WAIT_STEP)


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


def extract_company_code_from_name(file_name: str) -> str:
    name = file_name or ""
    match = re.search(r"\s-\s*(\d{1,4})\s-", name)
    if match:
        return str(int(match.group(1)))
    for part in name.split(" - "):
        if part.isdigit():
            return str(int(part))
    return ""

def resolve_year_month(output_root: Path, year: str | None, month: str | None) -> tuple[str, str]:
    if year and month:
        return year, month.zfill(2)

    years = []
    if output_root.exists():
        for entry in output_root.iterdir():
            if entry.is_dir() and entry.name.isdigit() and len(entry.name) == 4:
                years.append(entry.name)
    if not years:
        raise FileNotFoundError(f"Nenhuma pasta de ano encontrada em: {output_root}")
    years.sort()
    year = year or years[-1]

    months = []
    year_dir = output_root / year
    if year_dir.exists():
        for entry in year_dir.iterdir():
            if entry.is_dir() and entry.name.isdigit() and 1 <= len(entry.name) <= 2:
                months.append(entry.name.zfill(2))
    if not months:
        raise FileNotFoundError(f"Nenhuma pasta de mes encontrada em: {year_dir}")
    months.sort()
    month = month.zfill(2) if month else months[-1]
    return year, month


def collect_empresas_by_group(
    output_root: Path, year: str, month: str, grupos: list[str]
) -> dict[str, list[str]]:
    empresas_by_group: dict[str, set[str]] = {}
    for grupo in grupos:
        empresas_by_group[grupo] = set()
        group_dir = output_root / year / month / str(grupo)
        if not group_dir.exists():
            print(f"[warn] grupo {grupo} nao encontrado em {group_dir}")
            continue
        for pdf in group_dir.rglob("*.pdf"):
            code = extract_company_code_from_name(pdf.name)
            if code and code.isdigit():
                empresas_by_group[grupo].add(str(int(code)))
    return {g: sorted(vals, key=lambda x: int(x)) for g, vals in empresas_by_group.items()}


def col_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(r + ord("A")) + s
    return s


def write_xlsx_columns(path: Path, headers: list[str], columns: list[list[str]]) -> None:
    import zipfile
    from xml.sax.saxutils import escape

    if not headers or not columns or len(headers) != len(columns):
        raise ValueError("Headers e colunas precisam ter o mesmo tamanho.")

    def cell_inline(ref: str, value: str) -> str:
        return f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'

    max_rows = max((len(c) for c in columns), default=0)
    rows_xml = []

    header_cells = []
    for col_idx, header in enumerate(headers, 1):
        header_cells.append(cell_inline(f"{col_letter(col_idx)}1", str(header)))
    rows_xml.append(f'<row r="1">{"".join(header_cells)}</row>')

    for r in range(1, max_rows + 1):
        row_num = r + 1
        row_cells = []
        for col_idx, col_vals in enumerate(columns, 1):
            if r - 1 >= len(col_vals):
                continue
            val = str(col_vals[r - 1]).strip()
            if not val:
                continue
            row_cells.append(cell_inline(f"{col_letter(col_idx)}{row_num}", val))
        if row_cells:
            rows_xml.append(f'<row r="{row_num}">{"".join(row_cells)}</row>')

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

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def run_post_periodicos(coords, empresas):
    maybe_handle_error_dialog()
    do_click(coords["limpar"])
    time.sleep(0.1)

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

        maybe_handle_error_dialog()
        do_click(coords["limpar"])
        time.sleep(0.1)


def run_group_flow(coords, agrupamento: str, empresas: list[str], skip_gerar_periodicos: bool):
    if not empresas:
        print("[info] sem empresas encontradas")
        return

    alt_sequence("e", "m", "g")
    time.sleep(WAIT_SHORT)

    mes_ano = datetime.now().strftime("%m/%Y")
    do_type(coords["mes_ano"], mes_ano)
    time.sleep(WAIT_MED)

    do_click(coords["marcar_por_grupo"])
    time.sleep(WAIT_SHORT)

    do_type(coords["agrupamento"], str(agrupamento))
    time.sleep(WAIT_SHORT)

    do_click(coords["selecionar"])
    wait_or_skip(WAIT_SELECT, "selecao")

    do_click(coords["marcar_todas"])
    time.sleep(WAIT_MED)

    if skip_gerar_periodicos:
        print("[info] pular gerar periodicos")
    else:
        do_click(coords["gerar_periodicos"])
        wait_or_skip(WAIT_1_HOUR, "1 hora")

    run_post_periodicos(coords, empresas)


def main(
    skip_gerar_periodicos=False,
    coords_file=None,
    agrupamento=None,
    output_root=None,
    excel_dir=None,
    year=None,
    month=None,
):
    coords_path = Path(coords_file) if coords_file else COORDS_PATH
    coords = load_coords(coords_path)

    output_root = Path(output_root) if output_root else OUTPUT_ROOT
    agrupamento = agrupamento or DEFAULT_AGRUPAMENTO
    headers = [p.strip() for p in agrupamento.split(";") if p.strip()]
    year, month = resolve_year_month(output_root, year, month)
    empresas_by_group = collect_empresas_by_group(output_root, year, month, headers)
    empresas = sorted({v for vals in empresas_by_group.values() for v in vals}, key=lambda x: int(x))
    if not empresas:
        raise RuntimeError(f"Nenhuma empresa encontrada em: {output_root}\\{year}\\{month}")

    excel_dir = Path(excel_dir) if excel_dir else EXCEL_DIR
    excel_path = excel_dir / f"empresas_{year}{month}.xlsx"
    columns = [empresas_by_group.get(h, []) for h in headers]
    write_xlsx_columns(excel_path, headers=headers, columns=columns)
    print(f"[ok] excel salvo: {excel_path}")

    print("[info] foco em 3s")
    time.sleep(3)

    focus_folha_pagamento()
    time.sleep(WAIT_SHORT)

    hoje = datetime.now()
    data_sistema = date(hoje.year, hoje.month, DATA_DIA_FIXO).strftime("%d/%m/%Y")
    do_type(coords["data_sistema"], data_sistema)
    pag.press("enter")
    pag.press("enter")
    time.sleep(WAIT_MED)

    run_group_flow(coords, str(agrupamento), empresas, skip_gerar_periodicos=skip_gerar_periodicos)


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
        "--coords-file",
        default=None,
        help="Arquivo de coordenadas (.json). Padrao: arquivo com o mesmo nome do script.",
    )
    ap.add_argument(
        "--output-root",
        default=None,
        help="Raiz das pastas AUTOMATIZADO/1a PARTE.",
    )
    ap.add_argument(
        "--excel-dir",
        default=None,
        help="Pasta para salvar o Excel (padrao: Grupos por mes Excel).",
    )
    ap.add_argument(
        "--year",
        default=None,
        help="Ano (YYYY) para buscar a pasta YYYYMM.",
    )
    ap.add_argument(
        "--month",
        default=None,
        help="Mes (MM) para buscar a pasta YYYYMM.",
    )
    ap.add_argument(
        "--agrupamento",
        default=None,
        help="Texto do agrupamento (padrao: 6;13;14).",
    )
    args = ap.parse_args()

    skip_gerar = args.skip_gerar_periodicos or args.skip_emitir
    main(
        skip_gerar_periodicos=skip_gerar,
        coords_file=args.coords_file,
        agrupamento=args.agrupamento,
        output_root=args.output_root,
        excel_dir=args.excel_dir,
        year=args.year,
        month=args.month,
    )
