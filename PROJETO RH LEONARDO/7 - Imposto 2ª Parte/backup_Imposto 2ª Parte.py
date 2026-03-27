import argparse
import csv
import json
import re
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path

import msvcrt
import pyautogui as pag
import uiautomation as uia
try:
    import pygetwindow as gw
except Exception:
    gw = None
try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

COORDS_PATH = Path(__file__).with_suffix(".json")
BASE_DIR = Path(__file__).resolve().parent

DEFAULT_AGRUPAMENTO_FECHAMENTO = "7"
DEFAULT_AGRUPAMENTO_PERIODICOS = "31;32;33;34;35;36;37"

WAIT_SHORT = 0.3
WAIT_MED = 1.0
WAIT_STEP = 0.15
WAIT_TINY = 0.05
WAIT_1_HOUR = 60 * 60

WAIT_DARF_AFTER_SELECIONAR = 60.0
WAIT_DARF_AFTER_CONSULTAR = 5 * 60
WAIT_DARF_AFTER_EMITIR = 5 * 60
WAIT_FGTS_AFTER_SELECIONAR = 60.0
WAIT_FGTS_AFTER_CONSULTAR = 5 * 60
WAIT_FGTS_AFTER_EMITIR = 5 * 60

pag.PAUSE = 0.05
pag.FAILSAFE = False

DARF_COORDS_PATH = BASE_DIR / "Imposto 2Âª Parte - DARF.json"
FGTS_COORDS_PATH = BASE_DIR / "Imposto 2Âª Parte - FGTS.json"
DARF_SAVE_DIR_BASE = BASE_DIR / "DARF"
FGTS_SAVE_DIR_BASE = BASE_DIR / "FGTS"
OUTPUT_ROOT_1A = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\1Âª PARTE")
OUTPUT_ROOT_2A = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\2Âª PARTE")
OUTPUT_ROOT = OUTPUT_ROOT_2A
PROLABORE_ROOT = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\PRO LABORE")
BRUTOS_ROOT = BASE_DIR / "Arquivos"
IDLE_SECONDS = 30
DARF_CLEAR_COORD = (1370, 338)
FGTS_CLEAR_COORD = (1370, 392)


def load_coords(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de coordenadas nao encontrado: {path}")
    with path.open("r", encoding="utf-8-sig") as f:
        data = json.load(f)

    if "mes_ano" not in data:
        if "mes_ano_periodicos" not in data or "mes_ano_fechamento" not in data:
            raise KeyError("Coordenada ausente no JSON: mes_ano (ou mes_ano_periodicos + mes_ano_fechamento)")

    if "agrupamento_fechamento" not in data and "agrupamento" in data:
        data["agrupamento_fechamento"] = data["agrupamento"]

    required = [
        # Preflow (lista de empresas)
        "imprimir_inicial",
        "grupos_inicio",
        "grupos_fim",
        "imprimir_final",
        "salvar_pdf",
        "diretorio",
        "opcao_1",
        "opcao_2",
        "ok",
        "fechar_impressao",
        # Periodicos
        "marcar_por_grupo",
        "agrupamento",
        "selecionar",
        "marcar_todas",
        "gerar_periodicos",
        # Fechamento
        "marcar_agrupado",
        "agrupamento_fechamento",
        "consultar",
        "possui_remuneracao",
        "possui_pagamento_remuneracao",
        "transmissao_automatica_dctfweb",
        "situacao",
        "fechamento",
        "gerar",
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

    for key in ["mes_ano", "mes_ano_periodicos", "mes_ano_fechamento", "s1260", "s1270"]:
        if key in data:
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


def right_click_at(coord):
    x, y = coord
    pag.moveTo(x, y, duration=0.1)
    pag.click(button="right")
    time.sleep(WAIT_STEP)


def do_click(coord):
    click_at(coord)


def do_type(coord, text: str, clear=True):
    click_at(coord)
    time.sleep(WAIT_SHORT)
    if clear:
        pag.hotkey("ctrl", "a")
        pag.press("backspace")
    pag.typewrite(text, interval=0.02)
    time.sleep(WAIT_STEP)


def do_type_mes_ano_strict(coord, mes_ano: str, press_enter: bool = True):
    # Campo de competencia costuma manter mascara/valor anterior; limpa e redigita.
    click_at(coord)
    time.sleep(WAIT_SHORT)
    for _ in range(2):
        pag.hotkey("ctrl", "a")
        time.sleep(0.05)
        pag.press("delete")
        pag.press("backspace", presses=10, interval=0.01)
        pag.press("home")
        pag.press("delete", presses=10, interval=0.01)
        time.sleep(0.05)
    pag.typewrite(mes_ano, interval=0.03)
    time.sleep(WAIT_STEP)
    # Reforco final para garantir substituicao completa.
    pag.hotkey("ctrl", "a")
    pag.typewrite(mes_ano, interval=0.03)
    if press_enter:
        pag.press("enter")
    time.sleep(WAIT_STEP)


def do_type_custom_clear(coord, text: str, backspaces: int = 0, deletes: int = 0):
    click_at(coord)
    time.sleep(WAIT_SHORT)
    if backspaces > 0:
        pag.press("backspace", presses=backspaces, interval=0.02)
    if deletes > 0:
        pag.press("delete", presses=deletes, interval=0.02)
    pag.typewrite(text, interval=0.02)
    time.sleep(WAIT_STEP)


def alt_sequence(*keys):
    pag.keyDown("alt")
    for k in keys:
        pag.press(k)
        time.sleep(0.05)
    pag.keyUp("alt")
    time.sleep(WAIT_STEP)


def wait_or_skip(total_seconds, label):
    if total_seconds <= 0:
        return
    print(f"[info] esperar {label} (Enter para pular)")
    step = 0.2
    waited = 0.0
    while waited < total_seconds:
        if msvcrt.kbhit():
            ch = msvcrt.getwch()
            if ch == "\r":
                print("[ok] espera pulada")
                return
        time.sleep(step)
        waited += step


def _maximize_window_uia(wnd):
    try:
        wnd.Maximize()
        return
    except Exception:
        pass
    try:
        wp = wnd.GetWindowPattern()
        if wp:
            # UIA: Normal=0, Maximized=1, Minimized=2
            wp.SetWindowVisualState(1)
            return
    except Exception:
        pass
    try:
        pag.hotkey("win", "up")
        time.sleep(0.2)
        pag.hotkey("win", "up")
    except Exception:
        pass


def _focus_folha_via_uia() -> bool:
    candidates = [
        uia.WindowControl(Name="Folha de Pagamento", ClassName="TfrmPrincipal"),
        uia.WindowControl(Name="Folha de Pagamento"),
        uia.WindowControl(RegexName=".*Folha de Pagamento.*"),
        uia.WindowControl(RegexName=".*Folha.*"),
    ]
    for wnd in candidates:
        if not wnd.Exists(1):
            continue
        try:
            wnd.SetFocus()
            try:
                wnd.SetActive()
            except Exception:
                pass
            _maximize_window_uia(wnd)
            time.sleep(0.1)
            return True
        except Exception:
            continue
    return False


def _focus_folha_via_title() -> bool:
    if gw is None:
        return False
    try:
        wins = gw.getWindowsWithTitle("Folha de Pagamento")
    except Exception:
        return False
    for w in wins:
        try:
            if getattr(w, "isMinimized", False):
                w.restore()
                time.sleep(0.1)
            w.activate()
            time.sleep(0.1)
            try:
                w.maximize()
            except Exception:
                pass
            return True
        except Exception:
            continue
    return False


def focus_folha_pagamento(retries: int = 6, wait_s: float = 0.3):
    for _ in range(max(1, retries)):
        if _focus_folha_via_uia():
            return True
        if _focus_folha_via_title():
            return True
        time.sleep(wait_s)
    return False


def previous_month(now: datetime) -> tuple[int, int]:
    year = now.year
    month = now.month - 1
    if month == 0:
        month = 12
        year -= 1
    return year, month


def format_previous_month(now: datetime) -> str:
    year, month = previous_month(now)
    return f"{month:02d}/{year}"


def parse_grupos(value: str) -> list[str]:
    if not value:
        return []
    parts = [p.strip() for p in re.split(r"[;,]", value) if p.strip()]
    return parts


def previous_month_year_month(now: datetime) -> tuple[str, str]:
    year, month = previous_month(now)
    return str(year), f"{month:02d}"


def current_mm_yyyy(now: datetime | None = None) -> str:
    now = now or datetime.now()
    return f"{now.month:02d}-{now.year}"


def click_at_offset(coord, y_offset: int = 0):
    x, y = coord
    y += y_offset
    pag.moveTo(x, y, duration=0.1)
    pag.click()


def click_once_offset(coord, y_offset: int = 0):
    x, y = coord
    y += y_offset
    pag.moveTo(x, y, duration=0.1)
    pag.mouseDown(button="left")
    time.sleep(0.01)
    pag.mouseUp(button="left")


def wait_file_stable(path: Path, timeout_s: float = 45.0, stable_for_s: float = 2.0) -> bool:
    t0 = time.time()
    last_size = None
    stable_since = None
    while (time.time() - t0) < timeout_s:
        if not path.exists():
            time.sleep(0.2)
            continue
        size = path.stat().st_size
        if last_size is None or size != last_size:
            last_size = size
            stable_since = time.time()
        else:
            if stable_since and (time.time() - stable_since) >= stable_for_s:
                return True
        time.sleep(0.4)
    return False


def sanitize_filename(name: str) -> str:
    invalid = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for ch in invalid:
        name = name.replace(ch, "-")
    name = " ".join(name.split())
    return name.strip()


def _cnpj_pattern() -> re.Pattern:
    return re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")


def _format_company_folder(name: str, cnpj: str) -> str:
    name = sanitize_filename(name)
    cnpj = sanitize_filename(cnpj)
    if not name:
        return cnpj
    if not cnpj:
        return name
    return f"{name} {cnpj}"


def extract_company_from_pdf(pdf_path: Path) -> str:
    if fitz is None:
        return ""
    try:
        doc = fitz.open(pdf_path)
        if doc.page_count < 1:
            return ""
        page = doc.load_page(0)
        text = page.get_text("text") or ""
        data = page.get_text("dict")
        doc.close()
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        cnpj_re = _cnpj_pattern()
        label_re = re.compile(r"razao\s+social", re.IGNORECASE)
        cnpj_any = _extract_cnpj(" ".join(lines))

        money_re = re.compile(r"\d{1,3}(\.\d{3})*,\d{2}$")
        simple_number_re = re.compile(r"^\d{4,}$")
        blacklist = {"periodo", "apuracao", "data", "vencimento", "numero", "documento", "pagar"}

        def clean_candidate(raw: str) -> str:
            s = cnpj_re.sub("", raw).strip(" -:\t")
            if not s:
                return ""
            lowered = s.lower()
            if any(token in lowered for token in blacklist):
                return ""
            if money_re.search(s) or simple_number_re.search(s):
                return ""
            if any(ch.isdigit() for ch in s):
                return ""
            letters = sum(ch.isalpha() for ch in s)
            if letters < 5:
                return ""
            return s

        label_line = None
        label_block = None
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                line_text = " ".join(span.get("text", "") for span in line.get("spans", [])).strip()
                if label_re.search(line_text):
                    label_line = line
                    label_block = block
                    break
            if label_block:
                break

        if label_block and label_line:
            label_y = label_line.get("bbox", [0, 0, 0, 0])[3]
            candidates = []
            for line in label_block.get("lines", []):
                if line.get("bbox", [0, 0, 0, 0])[1] <= label_y + 1:
                    continue
                line_text = " ".join(span.get("text", "") for span in line.get("spans", [])).strip()
                cleaned = clean_candidate(line_text)
                if cleaned:
                    candidates.append((line.get("bbox", [0, 0, 0, 0])[1], cleaned))
            if candidates:
                candidates.sort(key=lambda x: x[0])
                return _format_company_folder(candidates[0][1], cnpj_any)

        label_periodo_re = re.compile(r"periodo\s+de\s+apuracao", re.IGNORECASE)
        for line in lines:
            if not cnpj_re.search(line):
                continue
            m_periodo = label_periodo_re.search(line)
            if m_periodo:
                line = line[: m_periodo.start()].strip()
            m = re.search(rf"({_cnpj_pattern().pattern})\s+(.+)", line)
            if not m:
                continue
            cnpj = m.group(1)
            name_raw = m.group(2)
            name = clean_candidate(name_raw)
            if name:
                return _format_company_folder(name, cnpj)

        for idx, line in enumerate(lines):
            if label_re.search(line):
                for offset in (1, 2):
                    if idx + offset >= len(lines):
                        break
                    nxt = clean_candidate(lines[idx + offset])
                    if nxt:
                        return nxt
                candidate = line
                if ":" in line:
                    candidate = line.split(":", 1)[1]
                candidate = clean_candidate(candidate)
                if candidate:
                    return _format_company_folder(candidate, cnpj_any)
            if cnpj_re.search(line) and not label_re.search(line):
                candidate = clean_candidate(line)
                if candidate:
                    return _format_company_folder(candidate, cnpj_any)
        return ""
    except Exception:
        return ""


def _normalize_company_key(name: str) -> str:
    base = (name or "").upper()
    return "".join(ch for ch in base if ch.isalpha())


def _company_folder_map(year: str, month: str) -> dict[str, str]:
    base = PROLABORE_ROOT / year / month
    if not base.exists():
        return {}
    mapping = {}
    for p in base.iterdir():
        if p.is_dir():
            mapping[_normalize_company_key(p.name)] = p.name
    return mapping


def _company_folder_map_1a_parte(year: str, month: str, grupo: str) -> dict[str, str]:
    base = OUTPUT_ROOT_1A / year / month / str(grupo)
    if not base.exists():
        return {}
    mapping = {}
    for p in base.iterdir():
        if p.is_dir():
            mapping[_normalize_company_key(p.name)] = p.name
    return mapping


def _strip_leading_code(name: str) -> str:
    if not name:
        return ""
    m = re.match(r"^\s*\d+\s*[- ]\s*(.+)$", name)
    return m.group(1).strip() if m else name.strip()


def _extract_leading_code(name: str) -> str:
    if not name:
        return ""
    m = re.match(r"^\s*(\d+)\s*[- ]", name)
    if not m:
        return ""
    try:
        return str(int(m.group(1)))
    except Exception:
        return m.group(1)


def _extract_cnpj(text: str) -> str:
    if not text:
        return ""
    m = _cnpj_pattern().search(text)
    return m.group(0) if m else ""


def _company_from_path(pdf_path: Path, marker: str) -> str:
    parts = [p.upper() for p in pdf_path.parts]
    try:
        idx = parts.index(marker)
    except ValueError:
        return ""
    if idx + 1 >= len(pdf_path.parts):
        return ""
    # Espera: .../<MARKER>/<mm-yyyy>/<grupo>/<empresa>/...
    date_re = re.compile(r"^\d{2}-\d{4}$")
    part1 = pdf_path.parts[idx + 1]
    if date_re.match(part1):
        if idx + 3 < len(pdf_path.parts) and pdf_path.parts[idx + 2].isdigit():
            return sanitize_filename(pdf_path.parts[idx + 3])
        return ""
    if idx + 2 < len(pdf_path.parts) and pdf_path.parts[idx + 1].isdigit():
        return sanitize_filename(pdf_path.parts[idx + 2])
    return sanitize_filename(part1)


def _extract_code_from_1a_parte_filename(name: str) -> str:
    m = re.search(r"\s-\s*(\d+)\s-\s*folha", name, re.IGNORECASE)
    if not m:
        return ""
    try:
        return str(int(m.group(1)))
    except Exception:
        return m.group(1)


def _code_folder_map_1a_parte(year: str, month: str, grupo: str) -> dict[str, str]:
    base = OUTPUT_ROOT_1A / year / month / str(grupo)
    if not base.exists():
        return {}
    mapping = {}
    for pdf in base.rglob("*.pdf"):
        code = _extract_code_from_1a_parte_filename(pdf.name)
        if not code:
            continue
        parent = pdf.parent
        if parent.is_dir():
            mapping.setdefault(code, parent.name)
    return mapping


def move_pdf_to_output(
    pdf_path: Path,
    grupo: str,
    marker: str,
    suffix: str,
    output_root: Path = OUTPUT_ROOT,
    year: str | None = None,
    month: str | None = None,
    code_map_1a: dict[str, str] | None = None,
    folder_map_1a: dict[str, str] | None = None,
    folder_map_prolabore: dict[str, str] | None = None,
    simple_filename: str | None = None,
):
    now = datetime.now()
    year = year or str(now.year)
    month = month or f"{now.month:02d}"
    prev_year, prev_month = previous_month(now)
    prev_label = f"{prev_month:02d}-{prev_year}"

    company_from_folder = _company_from_path(pdf_path, marker)
    code_from_folder = _extract_leading_code(company_from_folder)
    company = _strip_leading_code(company_from_folder)
    if not company or company.isdigit():
        company = extract_company_from_pdf(pdf_path)
    if not company:
        company = sanitize_filename(pdf_path.stem)
    if not company:
        print(f"[pdf] nome nao encontrado na pasta-mae: {pdf_path}")
        company = "SEM_NOME"
    code_map_1a = code_map_1a or _code_folder_map_1a_parte(year, month, grupo)
    if code_from_folder:
        mapped_by_code = code_map_1a.get(code_from_folder)
        if mapped_by_code:
            print(f"[pdf] mapeado via codigo 1Âª parte: '{code_from_folder}' -> '{mapped_by_code}'")
            company = mapped_by_code
    folder_map_1a = folder_map_1a or _company_folder_map_1a_parte(year, month, grupo)
    mapped = folder_map_1a.get(_normalize_company_key(company))
    if mapped:
        print(f"[pdf] mapeado via 1Âª parte: '{company}' -> '{mapped}'")
        company = mapped
    folder_map_prolabore = folder_map_prolabore or _company_folder_map(year, month)
    mapped = folder_map_prolabore.get(_normalize_company_key(company))
    if mapped:
        print(f"[pdf] mapeado via pro-labore: '{company}' -> '{mapped}'")
        company = mapped
    else:
        print(f"[pdf] sem correspondencia no pro-labore, criando sem codigo: '{company}'")
    cnpj = _extract_cnpj(company)
    if cnpj:
        name_only = _cnpj_pattern().sub("", company)
        name_only = sanitize_filename(name_only)
        company = _format_company_folder(name_only, cnpj)

    dest_dir = output_root / year / month / str(grupo) / company
    dest_dir.mkdir(parents=True, exist_ok=True)
    if simple_filename:
        dest_path = dest_dir / simple_filename
    else:
        dest_path = dest_dir / f"{company} - {suffix} {prev_label}.pdf"
    if dest_path.exists():
        stem = dest_path.stem
        ext = dest_path.suffix
        idx = 2
        while True:
            alt = dest_dir / f"{stem} ({idx}){ext}"
            if not alt.exists():
                dest_path = alt
                break
            idx += 1
        print(f"[pdf] arquivo ja existe, usando nome alternativo: {dest_path.name}")
    print(f"[pdf] origem: {pdf_path}")
    print(f"[pdf] destino: {dest_path}")
    shutil.copy2(str(pdf_path), str(dest_path))
    print(f"[pdf] copiado: {dest_path}")


def _competencia_from_path(pdf_path: Path) -> tuple[str, str] | None:
    # Procura um segmento YYYYMM (ex.: 202601) na arvore e usa como ano/mes do output.
    pat = re.compile(r"^\d{6}$")
    found = None
    for part in pdf_path.parts:
        if pat.match(part):
            found = part
    if not found:
        return None
    year = found[:4]
    month = found[4:6]
    if month < "01" or month > "12":
        return None
    return year, month


def _build_code_to_grupo_from_darf_brutos(darf_root: Path) -> dict[str, str]:
    mapping = {}
    if not darf_root.exists():
        return mapping
    for group_dir in darf_root.iterdir():
        if not group_dir.is_dir():
            continue
        grupo = group_dir.name.strip()
        if not grupo.isdigit():
            continue
        for company_dir in group_dir.iterdir():
            if not company_dir.is_dir():
                continue
            code = _extract_leading_code(company_dir.name)
            if code and code not in mapping:
                mapping[code] = grupo
    return mapping


def _build_code_to_company_from_darf_brutos(darf_root: Path) -> dict[str, str]:
    mapping = {}
    if not darf_root.exists():
        return mapping
    for group_dir in darf_root.iterdir():
        if not group_dir.is_dir():
            continue
        grupo = group_dir.name.strip()
        if not grupo.isdigit():
            continue
        for company_dir in group_dir.iterdir():
            if not company_dir.is_dir():
                continue
            code = _extract_leading_code(company_dir.name)
            if not code:
                continue
            company = sanitize_filename(_strip_leading_code(company_dir.name))
            if company and code not in mapping:
                mapping[code] = company
    return mapping


def organizar_brutos(
    brutos_root: Path,
    tipos: list[str] | None = None,
    limite: int | None = None,
):
    tipos = tipos or ["DARF", "FGTS"]
    brutos_root = Path(brutos_root)
    darf_root = brutos_root / "DARF"
    darf_code_to_grupo = _build_code_to_grupo_from_darf_brutos(darf_root)
    darf_code_to_company = _build_code_to_company_from_darf_brutos(darf_root)

    def _unique_path(path: Path) -> Path:
        if not path.exists():
            return path
        stem = path.stem
        ext = path.suffix
        idx = 2
        while True:
            alt = path.with_name(f"{stem} ({idx}){ext}")
            if not alt.exists():
                return alt
            idx += 1

    def _copy_to_automatizado(
        pdf_path: Path,
        grupo: str,
        company: str,
        year: str,
        month: str,
        tipo: str,
    ):
        company = sanitize_filename(company) or "SEM_NOME"
        dest_dir = OUTPUT_ROOT / year / month / str(grupo) / company
        dest_dir.mkdir(parents=True, exist_ok=True)
        # Nome final deve bater com o nome da pasta da empresa.
        # DARF vira "INSS"; FGTS mantem o tipo no nome.
        label = "INSS" if tipo.upper() == "DARF" else tipo.upper()
        dest_filename = f"{dest_dir.name} {label}.pdf"
        dest_path = _unique_path(dest_dir / dest_filename)
        print(f"[pdf] origem: {pdf_path}")
        print(f"[pdf] destino: {dest_path}")
        shutil.copy2(str(pdf_path), str(dest_path))
        print(f"[pdf] copiado: {dest_path}")

    def _iter_tipos(t: list[str]) -> list[str]:
        out = []
        for item in t:
            s = (item or "").strip().upper()
            if not s:
                continue
            out.append(s)
        return out or ["DARF", "FGTS"]

    tipos = _iter_tipos(tipos)

    copied = 0
    for tipo in tipos:
        tipo_root = brutos_root / tipo
        if not tipo_root.exists():
            print(f"[aviso] pasta de brutos nao encontrada: {tipo_root}")
            continue

        # Caso 1: ja existe /<TIPO>/<GRUPO>/...
        group_dirs = [d for d in tipo_root.iterdir() if d.is_dir() and d.name.isdigit()]
        if group_dirs:
            for group_dir in sorted(group_dirs, key=lambda p: int(p.name)):
                grupo = group_dir.name
                pdfs = sorted(group_dir.rglob("*.pdf"))
                if not pdfs:
                    continue
                print(f"[info] brutos {tipo} grupo {grupo}: {len(pdfs)} PDF(s)")
                for pdf in pdfs:
                    comp = _competencia_from_path(pdf)
                    if comp:
                        year, month = comp
                    else:
                        now = datetime.now()
                        py, pm = previous_month(now)
                        year, month = str(py), f"{pm:02d}"
                    company_from_folder = _company_from_path(pdf, tipo)
                    code = _extract_leading_code(company_from_folder)
                    company = (
                        darf_code_to_company.get(code)
                        if (tipo == "FGTS" and code)
                        else sanitize_filename(_strip_leading_code(company_from_folder))
                    )
                    _copy_to_automatizado(
                        pdf,
                        grupo=grupo,
                        company=company,
                        year=year,
                        month=month,
                        tipo=tipo,
                    )
                    copied += 1
                    if limite is not None and copied >= limite:
                        return
            continue

        # Caso 2: /<TIPO>/<EMPRESA>/... (sem grupo) -> tenta descobrir pelo codigo.
        pdfs = sorted(tipo_root.rglob("*.pdf"))
        if not pdfs:
            continue
        print(f"[info] brutos {tipo} sem grupo: {len(pdfs)} PDF(s)")
        for pdf in pdfs:
            rel = pdf.relative_to(tipo_root)
            company_folder = rel.parts[0] if rel.parts else ""
            code = _extract_leading_code(company_folder)
            comp = _competencia_from_path(pdf)
            if comp:
                year, month = comp
            else:
                now = datetime.now()
                py, pm = previous_month(now)
                year, month = str(py), f"{pm:02d}"

            grupo = ""
            if code:
                grupo = darf_code_to_grupo.get(code) or ""
            if not grupo:
                print(f"[aviso] nao foi possivel descobrir grupo para: {pdf}")
                continue

            company = darf_code_to_company.get(code) if code else sanitize_filename(_strip_leading_code(company_folder))
            _copy_to_automatizado(
                pdf,
                grupo=grupo,
                company=company,
                year=year,
                month=month,
                tipo=tipo,
            )
            copied += 1
            if limite is not None and copied >= limite:
                return


def process_pdfs(root_dir: Path, grupo: str, marker: str, suffix: str, idle_s: int = IDLE_SECONDS):
    root_dir.mkdir(parents=True, exist_ok=True)
    seen = set()
    last_activity = time.time()
    had_activity = False
    last_scan = 0.0
    now = datetime.now()
    prev_year, prev_month = previous_month(now)
    year = str(prev_year)
    month = f"{prev_month:02d}"
    code_map_1a = _code_folder_map_1a_parte(year, month, grupo)
    folder_map_1a = _company_folder_map_1a_parte(year, month, grupo)
    folder_map_prolabore = _company_folder_map(year, month)
    while True:
        activity = False
        if (time.time() - last_scan) >= 2.0:
            pdfs = sorted(root_dir.rglob("*.pdf"))
            for pdf in pdfs:
                if pdf in seen:
                    continue
                if not wait_file_stable(pdf):
                    continue
                move_pdf_to_output(
                    pdf,
                    grupo=grupo,
                    marker=marker,
                    suffix=suffix,
                    year=year,
                    month=month,
                    code_map_1a=code_map_1a,
                    folder_map_1a=folder_map_1a,
                    folder_map_prolabore=folder_map_prolabore,
                )
                seen.add(pdf)
                had_activity = True
                activity = True
            last_scan = time.time()
        if activity:
            last_activity = time.time()
        if had_activity and (time.time() - last_activity) >= idle_s:
            break
        time.sleep(0.5)


def load_coords_darf(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de coordenadas nao encontrado: {path}")
    with path.open("r", encoding="utf-8-sig") as f:
        data = json.load(f)
    required = [
        "centralizador",
        "modulos_extras",
        "controle_guias",
        "painel_controle_guias",
        "mes_ano",
        "agrupamento",
        "selecionar",
        "marcar_todas",
        "consultar",
        "emitir_darf",
        "compartilhar",
    ]
    for key in required:
        if key not in data:
            raise KeyError(f"Coordenada ausente no JSON (DARF): {key}")
    return data


def load_coords_fgts(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de coordenadas nao encontrado: {path}")
    with path.open("r", encoding="utf-8-sig") as f:
        data = json.load(f)
    required = [
        "centralizador",
        "modulos_extras",
        "controle_guias",
        "painel_controle_guias",
        "fgts",
        "mes_ano",
        "agrupamento",
        "selecionar",
        "marcar_todas",
        "consultar",
        "emitir_darf",
        "compartilhar",
    ]
    for key in required:
        if key not in data:
            raise KeyError(f"Coordenada ausente no JSON (FGTS): {key}")
    return data


def type_mes_ano_previous(coord):
    mes_ano = format_previous_month(datetime.now())
    do_type_mes_ano_strict(coord, mes_ano, press_enter=False)
    print(f"[info] competencia aplicada: {mes_ano}")


def select_group_offset(coord, grupo: str):
    click_at_offset(coord, y_offset=-2)
    time.sleep(WAIT_SHORT)
    pag.press("backspace", presses=4, interval=0.02)
    pag.typewrite(str(grupo), interval=0.02)
    pag.press("enter")
    time.sleep(3.0)


def click_selecionar_pasta():
    wnd = uia.WindowControl(Name="SeleÃ§Ã£o de DiretÃ³rio")
    if not wnd.Exists(2):
        return False
    try:
        wnd.SetFocus()
    except Exception:
        pass
    try:
        for c in wnd.GetChildren():
            if c.ControlTypeName != "ButtonControl":
                continue
            name = getattr(c, "Name", "") or ""
            if name.strip().lower() == "selecionar pasta":
                try:
                    inv = c.GetInvokePattern()
                    if inv:
                        inv.Invoke()
                        return True
                except Exception:
                    pass
                try:
                    c.Click()
                    return True
                except Exception:
                    pass
    except Exception:
        return False
    return False


def share_to_folder_offset(coords: dict, target_dir: Path, focus_fn):
    target_dir.mkdir(parents=True, exist_ok=True)
    focus_fn()
    time.sleep(WAIT_SHORT)
    click_at_offset(coords["compartilhar"], y_offset=-2)
    time.sleep(WAIT_MED)
    pag.press("down")
    time.sleep(WAIT_MED)
    pag.press("down")
    time.sleep(WAIT_MED)
    pag.press("enter")
    time.sleep(WAIT_MED)
    pag.hotkey("ctrl", "l")
    time.sleep(WAIT_SHORT)
    pag.typewrite(str(target_dir), interval=0.02)
    pag.press("enter")
    time.sleep(WAIT_MED)
    if not click_selecionar_pasta():
        pag.press("enter")


def focus_gerenciador_sistemas():
    wnd = uia.WindowControl(Name="Gerenciador de Sistemas")
    if not wnd.Exists(2):
        return False
    try:
        wnd.SetFocus()
        return True
    except Exception:
        return False


def run_darf_flow(coords: dict, grupos: list[str], skip_emitir: bool, skip_organizar: bool):
    print("[info] iniciando fluxo DARF")
    focus_folha_pagamento()
    time.sleep(WAIT_SHORT)

    click_at_offset(coords["centralizador"], y_offset=-2)
    time.sleep(WAIT_SHORT)
    click_at_offset(coords["modulos_extras"], y_offset=-2)
    time.sleep(WAIT_SHORT)
    click_at_offset(coords["controle_guias"], y_offset=-2)
    time.sleep(WAIT_SHORT)
    click_at_offset(coords["painel_controle_guias"], y_offset=-2)
    time.sleep(WAIT_SHORT)

    type_mes_ano_previous(coords["mes_ano"])
    time.sleep(WAIT_MED)

    date_folder = current_mm_yyyy()
    for grupo in grupos:
        print(f"[grupo] DARF {grupo}")
        click_at_offset(DARF_CLEAR_COORD, y_offset=-2)
        time.sleep(WAIT_SHORT)
        select_group_offset(coords["agrupamento"], grupo)

        click_at_offset(coords["selecionar"], y_offset=-2)
        time.sleep(WAIT_DARF_AFTER_SELECIONAR)
        click_once_offset(coords["marcar_todas"], y_offset=-2)
        time.sleep(WAIT_SHORT)

        if not skip_emitir:
            click_at_offset(coords["consultar"], y_offset=-2)
            wait_or_skip(WAIT_DARF_AFTER_CONSULTAR, "consultar DARF")
            click_once_offset(coords["marcar_todas"], y_offset=-2)
            time.sleep(WAIT_SHORT)

            click_at_offset(coords["emitir_darf"], y_offset=-2)
            wait_or_skip(WAIT_DARF_AFTER_EMITIR, "emitir DARF")
            click_once_offset(coords["marcar_todas"], y_offset=-2)
            time.sleep(WAIT_SHORT)

        share_to_folder_offset(
            coords,
            DARF_SAVE_DIR_BASE / date_folder / str(grupo),
            focus_folha_pagamento,
        )
        if not skip_organizar:
            process_pdfs(
                DARF_SAVE_DIR_BASE / date_folder / str(grupo),
                grupo=grupo,
                marker="DARF",
                suffix="DARF",
            )


def run_fgts_flow(coords: dict, grupos: list[str], skip_emitir: bool):
    print("[info] iniciando fluxo FGTS")
    focus_gerenciador_sistemas()
    time.sleep(WAIT_SHORT)

    click_at_offset(coords["centralizador"], y_offset=-2)
    time.sleep(WAIT_SHORT)
    click_at_offset(coords["modulos_extras"], y_offset=-2)
    time.sleep(WAIT_SHORT)
    click_at_offset(coords["controle_guias"], y_offset=-2)
    time.sleep(WAIT_SHORT)
    click_at_offset(coords["painel_controle_guias"], y_offset=-2)
    time.sleep(WAIT_SHORT)
    time.sleep(2.0)
    click_at_offset(coords["fgts"], y_offset=-2)
    time.sleep(WAIT_SHORT)

    type_mes_ano_previous(coords["mes_ano"])
    time.sleep(WAIT_MED)

    date_folder = current_mm_yyyy()
    for grupo in grupos:
        print(f"[grupo] FGTS {grupo}")
        click_at_offset(FGTS_CLEAR_COORD, y_offset=-2)
        time.sleep(WAIT_SHORT)
        select_group_offset(coords["agrupamento"], grupo)

        click_at_offset(coords["selecionar"], y_offset=-2)
        time.sleep(WAIT_FGTS_AFTER_SELECIONAR)
        click_once_offset(coords["marcar_todas"], y_offset=-2)
        time.sleep(WAIT_SHORT)

        if not skip_emitir:
            click_at_offset(coords["emitir_darf"], y_offset=-2)
            wait_or_skip(WAIT_FGTS_AFTER_EMITIR, "emitir FGTS")
            click_once_offset(coords["marcar_todas"], y_offset=-2)
            time.sleep(WAIT_SHORT)

            click_at_offset(coords["consultar"], y_offset=-2)
            wait_or_skip(WAIT_FGTS_AFTER_CONSULTAR, "consultar FGTS")
            click_once_offset(coords["marcar_todas"], y_offset=-2)
            time.sleep(WAIT_SHORT)

        share_to_folder_offset(
            coords,
            FGTS_SAVE_DIR_BASE / date_folder / str(grupo),
            focus_gerenciador_sistemas,
        )
        process_pdfs(
            FGTS_SAVE_DIR_BASE / date_folder / str(grupo),
            grupo=grupo,
            marker="FGTS",
            suffix="FGTS",
        )


def load_empresas_csv(path: Path) -> list[str]:
    if not path.exists():
        return []
    empresas = []
    suffix = path.suffix.lower()

    if suffix == ".slk":
        # SLK: captura codigos da coluna X1 (linhas de empresas).
        with path.open("r", encoding="utf-8", errors="ignore") as f:
            for raw in f:
                line = raw.strip()
                if not line.startswith("C;") or ";X1;" not in line or ";K" not in line:
                    continue
                value = line.split(";K", 1)[1].strip()
                if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
                    value = value[1:-1]
                code = value.strip()
                if not code.isdigit():
                    continue
                empresas.append(code)
        return empresas

    with path.open("r", encoding="utf-8", errors="ignore") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row:
                continue
            code = (row[0] or "").strip()
            if not code:
                continue
            empresas.append(code)
    return empresas


def ensure_coords(coords: dict, keys: list[str], label: str) -> None:
    missing = [k for k in keys if k not in coords]
    if missing:
        raise KeyError(f"Coordenadas ausentes para {label}: {', '.join(missing)}")


def menu_action(coord, keys, repeats=2):
    for _ in range(repeats):
        right_click_at(coord)
        time.sleep(0.5)
        for k in keys:
            pag.press(k)
            time.sleep(0.5)


def find_latest_pdf(output_dir: Path, after_ts: float) -> Path | None:
    if not output_dir.exists():
        return None
    pdfs = [p for p in output_dir.glob("*.pdf") if p.stat().st_mtime >= after_ts]
    if not pdfs:
        return None
    return max(pdfs, key=lambda p: p.stat().st_mtime)


def extract_text_from_pdf(pdf_path: Path) -> str:
    try:
        import PyPDF2  # type: ignore
    except Exception:
        PyPDF2 = None

    if PyPDF2:
        reader = PyPDF2.PdfReader(str(pdf_path))
        texts = []
        for page in reader.pages:
            try:
                texts.append(page.extract_text() or "")
            except Exception:
                texts.append("")
        return "\n".join(texts)

    if shutil.which("pdftotext"):
        tmp_txt = pdf_path.with_suffix(".txt")
        subprocess.run(
            ["pdftotext", str(pdf_path), str(tmp_txt)],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        if tmp_txt.exists():
            return tmp_txt.read_text(encoding="utf-8", errors="ignore")

    raise RuntimeError("Nao foi possivel extrair texto do PDF (PyPDF2/pdftotext ausentes).")


def extract_company_codes(text: str) -> list[str]:
    lines = text.splitlines()
    codes = []
    seen = set()
    pattern = re.compile(r"\s-\s*(\d{1,4})\s-")
    for line in lines:
        m = pattern.search(line)
        if not m:
            continue
        code = m.group(1)
        if code not in seen:
            seen.add(code)
            codes.append(code)
    if codes:
        return codes

    # Fallback: qualquer numero de 1-4 digitos
    for line in lines:
        for m in re.finditer(r"\b(\d{1,4})\b", line):
            code = m.group(1)
            if code not in seen:
                seen.add(code)
                codes.append(code)
    return codes


def write_empresas_csv(path: Path, codes: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for code in codes:
            writer.writerow([code])


def run_preflow(coords, grupos_inicio: str, grupos_fim: str, output_dir: Path):
    alt_sequence("d", "u")
    time.sleep(WAIT_MED)

    do_click(coords["imprimir_inicial"])
    time.sleep(WAIT_SHORT)

    do_type(coords["grupos_inicio"], grupos_inicio)
    time.sleep(WAIT_SHORT)
    do_type_custom_clear(coords["grupos_fim"], grupos_fim, backspaces=2, deletes=2)
    time.sleep(WAIT_SHORT)

    do_click(coords["imprimir_final"])
    time.sleep(WAIT_MED)

    start_ts = time.time()
    do_click(coords["salvar_pdf"])
    time.sleep(WAIT_MED)

    do_click(coords["diretorio"])
    time.sleep(WAIT_SHORT)
    pag.hotkey("ctrl", "a")
    time.sleep(WAIT_SHORT)
    pag.typewrite(str(output_dir), interval=0.01)
    time.sleep(WAIT_SHORT)

    do_click(coords["opcao_1"])
    time.sleep(WAIT_SHORT)
    do_click(coords["opcao_2"])
    time.sleep(WAIT_SHORT)

    do_click(coords["ok"])
    time.sleep(WAIT_SHORT)
    pag.press("enter")
    time.sleep(WAIT_MED)

    do_click(coords["fechar_impressao"])
    time.sleep(WAIT_MED)

    latest_pdf = find_latest_pdf(output_dir, after_ts=start_ts)
    if not latest_pdf:
        print("[aviso] nenhum PDF encontrado para gerar CSV")
        return

    try:
        text = extract_text_from_pdf(latest_pdf)
    except Exception as exc:
        print(f"[aviso] falha ao ler PDF: {exc}")
        return

    codes = extract_company_codes(text)
    if not codes:
        print("[aviso] nenhum codigo encontrado no PDF")
        return

    csv_path = output_dir / "empresas.csv"
    write_empresas_csv(csv_path, codes)
    print(f"[ok] empresas.csv gerado: {csv_path} ({len(codes)})")


def run_periodicos(
    coords,
    agrupamento: str,
    skip_gerar: bool,
    wait_gerar: float,
    wait_selecao: float,
    wait_antes_limpar: float,
    wait_depois_limpar: float,
    empresas: list[str] | None,
):
    # Ordem obrigatoria do fluxo:
    # 1) atalho abre tela, 2) digita competencia, 3) marcar por grupo, 4) digitar agrupamento.
    if not focus_folha_pagamento():
        raise RuntimeError("Nao foi possivel focar a janela 'Folha de Pagamento' em periodicos.")
    time.sleep(1.0)

    alt_sequence("e", "m", "g")
    time.sleep(WAIT_MED + 0.3)

    # Reforca foco apos abrir via bind, antes de preencher campos.
    if not focus_folha_pagamento(retries=2, wait_s=0.2):
        raise RuntimeError("Nao foi possivel manter foco na janela 'Folha de Pagamento' apos bind de periodicos.")
    time.sleep(WAIT_SHORT)

    mes_ano = format_previous_month(datetime.now())

    mes_coord = coords.get("mes_ano_periodicos", coords.get("mes_ano"))
    do_type_mes_ano_strict(mes_coord, mes_ano, press_enter=True)
    print(f"[info] competencia periodicos: {mes_ano}")
    time.sleep(WAIT_MED)

    do_click(coords["marcar_por_grupo"])
    time.sleep(WAIT_SHORT)

    do_type(coords["agrupamento"], str(agrupamento))
    time.sleep(WAIT_SHORT)

    do_click(coords["selecionar"])
    wait_or_skip(wait_selecao, "selecao")

    do_click(coords["marcar_todas"])
    time.sleep(WAIT_MED)

    if skip_gerar:
        print("[info] pular gerar periodicos")
    else:
        do_click(coords["gerar_periodicos"])
        wait_or_skip(wait_gerar, "1 hora")
        wait_or_skip(wait_antes_limpar, "antes limpar")

    if empresas:
        run_conferencia_periodicos(coords, empresas)
        wait_or_skip(wait_depois_limpar, "depois limpar")
    elif "limpar" in coords:
        do_click(coords["limpar"])
        wait_or_skip(wait_depois_limpar, "depois limpar")
    else:
        print("[aviso] coordenada limpar ausente: pular limpar")


def run_conferencia_periodicos(coords, empresas: list[str]) -> None:
    if not empresas:
        print("[info] sem empresas para conferencia")
        return
    ensure_coords(
        coords,
        [
            "campo_empresa",
            "aba_com_erro",
            "marcar_todas_erro",
            "reenviar",
            "confirmar_reenviar",
            "limpar",
        ],
        "conferencia de erros",
    )

    do_click(coords["limpar"])
    time.sleep(WAIT_TINY)

    for empresa in empresas:
        do_click(coords["marcar_por_grupo"])
        do_type(coords["campo_empresa"], empresa)

        do_click(coords["selecionar"])
        time.sleep(1.0)

        do_click(coords["aba_com_erro"])
        time.sleep(WAIT_TINY)

        do_click(coords["marcar_todas_erro"])
        time.sleep(WAIT_TINY)
        do_click(coords["reenviar"])

        time.sleep(WAIT_SHORT)
        do_click(coords["confirmar_reenviar"])
        time.sleep(WAIT_SHORT)

        do_click(coords["limpar"])
        time.sleep(WAIT_TINY)


def run_fechamento(
    coords,
    agrupamento: str,
    wait_antes_consultar: float,
    wait_consulta: float,
    wait_gerar: float,
    wait_antes_marcar: float,
    skip_consulta: bool,
):
    if not focus_folha_pagamento():
        raise RuntimeError("Nao foi possivel focar a janela 'Folha de Pagamento' no fechamento.")

    alt_sequence("e", "f")
    time.sleep(WAIT_MED)

    mes_ano = format_previous_month(datetime.now())

    mes_coord = coords.get("mes_ano_fechamento", coords.get("mes_ano"))
    do_type_mes_ano_strict(mes_coord, mes_ano, press_enter=True)
    print(f"[info] competencia fechamento: {mes_ano}")
    time.sleep(WAIT_MED)

    do_click(coords["marcar_agrupado"])
    time.sleep(WAIT_SHORT)

    do_type(coords["agrupamento_fechamento"], str(agrupamento))
    pag.press("enter")
    time.sleep(WAIT_SHORT)

    if skip_consulta:
        print("[info] pular consultar")
    else:
        wait_or_skip(wait_antes_consultar, "antes consultar")
        do_click(coords["consultar"])
        wait_or_skip(wait_consulta, "10 minutos")

    wait_or_skip(wait_antes_marcar, "antes marcar")

    # Ordem exigida: s1260, s1270, possui remuneracao, possui pagamento, trans aut, fechamento, situacao
    if "s1260" in coords:
        menu_action(coords["s1260"], keys=["down", "down", "right", "down", "enter"], repeats=2)
    else:
        print("[aviso] coordenada s1260 ausente: pular desmarcacao")
    if "s1270" in coords:
        menu_action(coords["s1270"], keys=["down", "down", "right", "down", "enter"], repeats=2)
    else:
        print("[aviso] coordenada s1270 ausente: pular desmarcacao")

    menu_action(coords["possui_remuneracao"], keys=["down", "down", "right", "enter"], repeats=2)
    menu_action(
        coords["possui_pagamento_remuneracao"],
        keys=["down", "down", "right", "enter"],
        repeats=2,
    )
    menu_action(
        coords["transmissao_automatica_dctfweb"],
        keys=["down", "down", "right", "enter"],
        repeats=2,
    )
    menu_action(coords["fechamento"], keys=["down", "down", "right", "down", "enter"], repeats=2)
    menu_action(coords["situacao"], keys=["down", "down", "right", "enter"], repeats=2)

    do_click(coords["gerar"])
    wait_or_skip(wait_gerar, "gerar")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--organizar-brutos",
        action="store_true",
        help="Apenas organiza PDFs ja baixados em BRUTOS_ROOT (sem automacao de UI).",
    )
    ap.add_argument(
        "--brutos-root",
        default=str(BRUTOS_ROOT),
        help="Pasta raiz dos brutos (padrao: ./Arquivos).",
    )
    ap.add_argument(
        "--organizar-tipos",
        default="DARF;FGTS",
        help="Tipos para organizar (ex: DARF;FGTS ou apenas DARF).",
    )
    ap.add_argument(
        "--organizar-limite",
        type=int,
        default=None,
        help="Limite de PDFs para copiar (teste).",
    )
    ap.add_argument(
        "--coords-file",
        default=None,
        help="Arquivo de coordenadas (.json). Padrao: arquivo com o mesmo nome do script.",
    )
    ap.add_argument(
        "--grupos-inicio",
        default="31",
        help="Valor para o campo grupos inicio (preflow).",
    )
    ap.add_argument(
        "--grupos-fim",
        default="37",
        help="Valor para o campo grupos fim (preflow).",
    )
    ap.add_argument(
        "--skip-preflow",
        action="store_true",
        help="Pula o preflow (lista de empresas).",
    )
    ap.add_argument(
        "--agrupamento",
        default=DEFAULT_AGRUPAMENTO_PERIODICOS,
        help="Texto do agrupamento para periodicos (padrao: 31;32;33;34;35;36;37).",
    )
    ap.add_argument(
        "--agrupamento-fechamento",
        default=DEFAULT_AGRUPAMENTO_FECHAMENTO,
        help="Texto do agrupamento para fechamento (padrao: 7).",
    )
    ap.add_argument(
        "--wait-gerar",
        type=float,
        default=60 * 60,
        help="Tempo de espera apos gerar periodicos.",
    )
    ap.add_argument(
        "--wait-selecao",
        type=float,
        default=10 * 60,
        help="Tempo de espera apos selecionar (antes de marcar todas).",
    )
    ap.add_argument(
        "--wait-antes-limpar",
        type=float,
        default=10 * 60,
        help="Tempo de espera apos gerar periodicos (antes de limpar).",
    )
    ap.add_argument(
        "--wait-depois-limpar",
        type=float,
        default=30 * 60,
        help="Tempo de espera apos limpar (antes de continuar).",
    )
    ap.add_argument(
        "--wait-antes-consultar",
        type=float,
        default=30 * 60,
        help="Tempo de espera antes de clicar em Consultar no fechamento.",
    )
    ap.add_argument(
        "--wait-consulta",
        type=float,
        default=10 * 60,
        help="Tempo de espera apos consultar no fechamento.",
    )
    ap.add_argument(
        "--wait-antes-marcar",
        type=float,
        default=10 * 60,
        help="Tempo de espera antes de marcar checkboxes no fechamento.",
    )
    ap.add_argument(
        "--wait-gerar-fechamento",
        type=float,
        default=10.0,
        help="Tempo de espera apos gerar no fechamento.",
    )
    ap.add_argument(
        "--skip-gerar-periodicos",
        action="store_true",
        help="Executa o fluxo sem clicar em Gerar Periodicos (modo teste).",
    )
    ap.add_argument(
        "--skip-periodicos",
        action="store_true",
        help="Pula o fluxo de periodicos.",
    )
    ap.add_argument(
        "--skip-fechamento",
        action="store_true",
        help="Pula o fluxo de fechamento.",
    )
    ap.add_argument(
        "--skip-consultar",
        action="store_true",
        help="Nao clica em Consultar no fechamento.",
    )
    ap.add_argument(
        "--skip-conferencia",
        action="store_true",
        help="Pula a conferencia automatica empresa por empresa.",
    )
    ap.add_argument(
        "--empresas-csv",
        default=str(BASE_DIR / "empresas.csv"),
        help="Arquivo CSV/SLK com codigos de empresa (padrao: empresas.csv na pasta do script).",
    )
    ap.add_argument(
        "--skip-darf-fgts",
        action="store_true",
        help="Pula o fluxo de DARF/FGTS apos o fechamento.",
    )
    ap.add_argument(
        "--darf-coords-file",
        default=str(DARF_COORDS_PATH),
        help="Arquivo de coordenadas do DARF.",
    )
    ap.add_argument(
        "--fgts-coords-file",
        default=str(FGTS_COORDS_PATH),
        help="Arquivo de coordenadas do FGTS.",
    )
    ap.add_argument(
        "--darf-grupos",
        default="31;32;33;34;35;36;37",
        help="Lista de grupos DARF (ex: 31;32;33;34;35;36;37).",
    )
    ap.add_argument(
        "--fgts-grupos",
        default="31;32;33;34;35;36;37",
        help="Lista de grupos FGTS (ex: 31;32;33;34;35;36;37).",
    )
    ap.add_argument(
        "--skip-emitir-darf",
        action="store_true",
        help="Nao clica em Emitir DARF.",
    )
    ap.add_argument(
        "--skip-emitir-fgts",
        action="store_true",
        help="Nao clica em Emitir FGTS.",
    )
    ap.add_argument(
        "--skip-organizar-darf",
        action="store_true",
        help="Nao organiza/renomeia PDFs de DARF na pasta automatizado.",
    )
    args = ap.parse_args()

    if args.organizar_brutos:
        tipos = [p.strip() for p in re.split(r"[;,]", args.organizar_tipos or "") if p.strip()]
        organizar_brutos(
            Path(args.brutos_root),
            tipos=tipos,
            limite=args.organizar_limite,
        )
        return

    coords_path = Path(args.coords_file) if args.coords_file else COORDS_PATH
    coords = load_coords(coords_path)

    grupos_inicio = (args.grupos_inicio or "").strip()
    grupos_fim = (args.grupos_fim or "").strip()
    if not args.skip_preflow:
        if not grupos_inicio:
            grupos_inicio = input("Grupos inicio (ex: 31): ").strip()
        if not grupos_fim:
            grupos_fim = input("Grupos fim (ex: 37): ").strip()

        if "-" in grupos_inicio and not grupos_fim:
            parts = [p.strip() for p in grupos_inicio.split("-", 1)]
            if len(parts) == 2 and parts[0] and parts[1]:
                grupos_inicio, grupos_fim = parts[0], parts[1]
    if not args.skip_preflow and (not grupos_inicio or not grupos_fim):
        raise SystemExit("Grupos inicio/fim nao informados.")

    agrupamento = (args.agrupamento or "").strip()
    if not agrupamento:
        raise SystemExit("Agrupamento nao informado (use --agrupamento).")

    now = datetime.now()
    if now.day != 7:
        print(f"[aviso] fluxo previsto para o dia 7 (hoje: {now:%d/%m/%Y}).")

    print("[info] foco em 3s")
    time.sleep(3)

    if not focus_folha_pagamento():
        raise RuntimeError("Nao foi possivel focar a janela 'Folha de Pagamento'.")
    time.sleep(WAIT_SHORT)

    if not args.skip_preflow:
        run_preflow(
            coords,
            grupos_inicio=grupos_inicio,
            grupos_fim=grupos_fim,
            output_dir=BASE_DIR,
        )

    if not args.skip_periodicos:
        empresas = [] if args.skip_conferencia else load_empresas_csv(Path(args.empresas_csv))
        if not args.skip_conferencia and not empresas:
            print(f"[aviso] arquivo de empresas vazio/ausente: {args.empresas_csv} (pular conferencia)")
        run_periodicos(
            coords,
            agrupamento=agrupamento,
            skip_gerar=args.skip_gerar_periodicos,
            wait_gerar=args.wait_gerar,
            wait_selecao=args.wait_selecao,
            wait_antes_limpar=args.wait_antes_limpar,
            wait_depois_limpar=args.wait_depois_limpar,
            empresas=empresas if not args.skip_conferencia else None,
        )

    if not args.skip_fechamento:
        run_fechamento(
            coords,
            agrupamento=args.agrupamento_fechamento,
            wait_antes_consultar=args.wait_antes_consultar,
            wait_consulta=args.wait_consulta,
            wait_gerar=args.wait_gerar_fechamento,
            wait_antes_marcar=args.wait_antes_marcar,
            skip_consulta=args.skip_consultar,
        )

    if not args.skip_darf_fgts:
        if not args.skip_fechamento:
            wait_or_skip(WAIT_1_HOUR, "1 hora antes DARF/FGTS")
        else:
            print("[info] fechamento pulado: iniciar DARF/FGTS sem espera de 1 hora")

        darf_grupos = parse_grupos(args.darf_grupos or "")
        if not darf_grupos:
            darf_input = input("Grupos DARF (ex: 6;13;14): ").strip()
            darf_grupos = parse_grupos(darf_input)
        if not darf_grupos:
            raise SystemExit("Grupos DARF nao informados.")

        fgts_grupos = parse_grupos(args.fgts_grupos or "")
        if not fgts_grupos:
            fgts_input = input("Grupos FGTS (ex: 6;13;14): ").strip()
            fgts_grupos = parse_grupos(fgts_input)
        if not fgts_grupos:
            raise SystemExit("Grupos FGTS nao informados.")

        darf_coords = load_coords_darf(Path(args.darf_coords_file))
        fgts_coords = load_coords_fgts(Path(args.fgts_coords_file))

        run_darf_flow(
            darf_coords,
            grupos=darf_grupos,
            skip_emitir=args.skip_emitir_darf,
            skip_organizar=args.skip_organizar_darf,
        )
        run_fgts_flow(
            fgts_coords,
            grupos=fgts_grupos,
            skip_emitir=args.skip_emitir_fgts,
        )



if __name__ == "__main__":
    main()

