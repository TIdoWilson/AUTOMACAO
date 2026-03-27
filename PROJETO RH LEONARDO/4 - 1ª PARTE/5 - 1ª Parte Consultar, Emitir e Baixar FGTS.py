import argparse
import json
import re
import time
import shutil
from datetime import datetime
from pathlib import Path

import msvcrt
import pyautogui as pag
import uiautomation as uia

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

BASE_DIR = Path(__file__).resolve().parent
COORDS_PATH = Path(__file__).with_suffix(".json")

WAIT_SHORT = 0.2
WAIT_MED = 0.8
WAIT_2_SEC = 2.0
WAIT_120_MIN = 120
WAIT_GROUP_CHANGE = 3.0
WAIT_AFTER_PAINEL = 2.0

pag.PAUSE = 0.5
pag.FAILSAFE = False

SAVE_DIR_BASE = BASE_DIR / "FGTS"
OUTPUT_ROOT = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\1ª PARTE")
PROLABORE_ROOT = Path(r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\PRO LABORE")
IDLE_SECONDS = 30
GRUPOS = ["6", "13", "14"]
CLEAR_COORD = (1370, 392)


def load_coords(path: Path) -> dict:
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
    y -= 2
    pag.moveTo(x, y, duration=0.1)
    pag.click()


def click_once(coord):
    x, y = coord
    y -= 2
    pag.moveTo(x, y, duration=0.1)
    pag.mouseDown(button="left")
    time.sleep(0.01)
    pag.mouseUp(button="left")


def type_mes_ano(coord):
    hoje = datetime.now()
    mes_ano = hoje.strftime("%m/%Y")
    click_at(coord)
    time.sleep(WAIT_SHORT)
    pag.press("backspace", presses=7, interval=0.02)
    pag.typewrite(mes_ano, interval=0.02)


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
    base = OUTPUT_ROOT / year / month / str(grupo)
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


def _company_from_path(pdf_path: Path) -> str:
    parts = [p.upper() for p in pdf_path.parts]
    try:
        idx = parts.index("FGTS")
    except ValueError:
        return ""
    if idx + 1 >= len(pdf_path.parts):
        return ""
    # Espera: ...\FGTS\<grupo>\<empresa>\...
    if idx + 2 < len(pdf_path.parts) and pdf_path.parts[idx + 1].isdigit():
        return sanitize_filename(pdf_path.parts[idx + 2])
    return sanitize_filename(pdf_path.parts[idx + 1])


def _extract_code_from_1a_parte_filename(name: str) -> str:
    m = re.search(r"\s-\s*(\d+)\s-\s*folha", name, re.IGNORECASE)
    if not m:
        return ""
    try:
        return str(int(m.group(1)))
    except Exception:
        return m.group(1)


def _code_folder_map_1a_parte(year: str, month: str, grupo: str) -> dict[str, str]:
    base = OUTPUT_ROOT / year / month / str(grupo)
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
    output_root: Path = OUTPUT_ROOT,
    year: str | None = None,
    month: str | None = None,
    code_map_1a: dict[str, str] | None = None,
    folder_map_1a: dict[str, str] | None = None,
    folder_map_prolabore: dict[str, str] | None = None,
):
    now = datetime.now()
    year = year or str(now.year)
    month = month or f"{now.month:02d}"

    company_from_folder = _company_from_path(pdf_path)
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
            print(f"[pdf] mapeado via codigo 1ª parte: '{code_from_folder}' -> '{mapped_by_code}'")
            company = mapped_by_code
    folder_map_1a = folder_map_1a or _company_folder_map_1a_parte(year, month, grupo)
    mapped = folder_map_1a.get(_normalize_company_key(company))
    if mapped:
        print(f"[pdf] mapeado via 1ª parte: '{company}' -> '{mapped}'")
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
    dest_path = dest_dir / f"{company} - FGTS.pdf"
    if dest_path.exists():
        idx = 2
        while True:
            alt = dest_dir / f"{company} - FGTS ({idx}).pdf"
            if not alt.exists():
                dest_path = alt
                break
            idx += 1
        print(f"[pdf] arquivo ja existe, usando nome alternativo: {dest_path.name}")
    print(f"[pdf] origem: {pdf_path}")
    print(f"[pdf] destino: {dest_path}")
    shutil.copy2(str(pdf_path), str(dest_path))
    print(f"[pdf] copiado: {dest_path}")


def process_darf_pdfs(root_dir: Path, grupo: str, idle_s: int = IDLE_SECONDS):
    root_dir.mkdir(parents=True, exist_ok=True)
    seen = set()
    last_activity = time.time()
    had_activity = False
    last_scan = 0.0
    now = datetime.now()
    year = str(now.year)
    month = f"{now.month:02d}"
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


def focus_gerenciador_sistemas():
    wnd = uia.WindowControl(Name="Gerenciador de Sistemas")
    if not wnd.Exists(2):
        return False
    try:
        wnd.SetFocus()
        return True
    except Exception:
        return False


def click_selecionar_pasta():
    wnd = None
    for title in ("Seleção de Diretório", "Selecao de Diretorio", "SeleÃ‡ÃµÃ‡Å“o de DiretÃ‡Ã¼rio"):
        test = uia.WindowControl(Name=title)
        if test.Exists(1):
            wnd = test
            break
    if wnd is None:
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


def select_group(coord, grupo: str):
    click_at(coord)
    time.sleep(WAIT_SHORT)
    pag.press("backspace", presses=4, interval=0.02)
    pag.typewrite(str(grupo), interval=0.02)
    pag.press("enter")
    time.sleep(WAIT_GROUP_CHANGE)


def share_to_folder(coords: dict, target_dir: Path):
    target_dir.mkdir(parents=True, exist_ok=True)
    focus_gerenciador_sistemas()
    time.sleep(WAIT_SHORT)
    click_at(coords["compartilhar"])
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


def main(
    skip_emitir=False,
    post_only=False,
    share_only=False,
    stop_after_marcar=False,
    grupo_only: str | None = None,
):
    coords = load_coords(COORDS_PATH)

    if post_only:
        grupos = [grupo_only] if grupo_only else GRUPOS
        for grupo in grupos:
            process_darf_pdfs(SAVE_DIR_BASE / str(grupo), grupo=grupo)
        return

    print("[info] foco em 3s")
    time.sleep(3)

    focus_gerenciador_sistemas()
    time.sleep(WAIT_SHORT)

    click_at(coords["centralizador"])
    time.sleep(WAIT_SHORT)

    click_at(coords["modulos_extras"])
    time.sleep(WAIT_SHORT)

    click_at(coords["controle_guias"])
    time.sleep(WAIT_SHORT)

    click_at(coords["painel_controle_guias"])
    time.sleep(WAIT_SHORT)
    time.sleep(WAIT_AFTER_PAINEL)
    click_at(coords["fgts"])
    time.sleep(WAIT_SHORT)

    type_mes_ano(coords["mes_ano"])
    time.sleep(WAIT_MED)

    grupos = [grupo_only] if grupo_only else GRUPOS
    for grupo in grupos:
        print(f"[grupo] iniciando {grupo}")
        click_at(CLEAR_COORD)
        time.sleep(WAIT_SHORT)
        select_group(coords["agrupamento"], grupo)

        click_at(coords["selecionar"])
        time.sleep(WAIT_SHORT)
        click_once(coords["marcar_todas"])
        time.sleep(WAIT_SHORT)

        if stop_after_marcar:
            continue

        if not share_only:
            time.sleep(WAIT_2_SEC)
            click_at(coords["consultar"])
            wait_or_skip(WAIT_120_MIN, "2 minutos")

            click_once(coords["marcar_todas"])
            time.sleep(WAIT_SHORT)

            if skip_emitir:
                print("[info] pular emitir DARF")
            else:
                click_at(coords["emitir_darf"])
                wait_or_skip(WAIT_120_MIN, "2 minutos")

            click_once(coords["marcar_todas"])
            time.sleep(WAIT_SHORT)

        share_to_folder(coords, SAVE_DIR_BASE / str(grupo))
        process_darf_pdfs(SAVE_DIR_BASE / str(grupo), grupo=grupo)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--skip-emitir",
        action="store_true",
        help="Executa o fluxo sem clicar em Emitir DARF.",
    )
    ap.add_argument(
        "--share-only",
        action="store_true",
        help="Somente compartilha/baixa os PDFs (sem consultar nem emitir).",
    )
    ap.add_argument(
        "--stop-after-marcar",
        action="store_true",
        help="Para apos marcar todas (sem consultar, emitir ou compartilhar).",
    )
    ap.add_argument(
        "--post-only",
        action="store_true",
        help="Executa apenas a organizacao/renomeacao dos PDFs.",
    )
    ap.add_argument(
        "--grupo",
        help="Processa apenas um agrupamento especifico (ex: 6, 13, 14).",
    )
    args = ap.parse_args()
    main(
        skip_emitir=args.skip_emitir,
        post_only=args.post_only,
        share_only=args.share_only,
        stop_after_marcar=args.stop_after_marcar,
        grupo_only=args.grupo,
    )

