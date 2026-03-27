import os
import re
import shutil
import subprocess
import sys
import time
import unicodedata
from datetime import datetime

# pyautogui so e necessario no modo de automacao (lazy import).
pag = None


# =========================
# CONFIG (100% COORDENADAS)
# =========================
# Fluxo (conforme solicitado):
# 1) Focar tela: 2 cliques em 1470 167
# 2) ALT M E
# 3) Sem clicar, digitar nessa ordem:
#    1
#    MM/YY (mes/ano atual)
#    1
#    30
# 4) Espera 5s
# 5) Fluxos com mouse 2 + teclas (delay 0.05 entre cada botao)
# 6) Clicar em Parametro / Relatorios mensais / (705,363 + backspace*2 + 1) / Salvar por empresa / Digitar caminho
# 7) ALT P

FOCUS_X = 1470
FOCUS_Y = 167
FOCUS_CLICKS = 2

MOUSE2_SEQS = [
    # (x, y, keys)
    (1175, 321, ["down", "down", "right", "enter"]),
    (1233, 323, ["down", "down", "right", "enter"]),
    (1286, 318, ["down", "down", "right", "down", "enter"]),
]
MOUSE2_LEFT_SHIFT_PX = 5

CLICK_PARAMETRO = (520, 292)
CLICK_RELATORIOS_MENSAIS = (683, 317)
CLICK_CAMPO_ANTES_SALVAR_POR_EMPRESA = (705, 363)
CLICK_SALVAR_ARQ_POR_EMPRESA = (597, 782)
CLICK_CAMINHO_DIRETORIO = (1010, 804)

DIRETORIO_ARQUIVOS = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\6 - Adiantamento\Arquivos"

# Organizador (modo separado, sem automacao de UI)
# Estrutura destino:
# AUTOMATIZADO\ADIANTAMENTO\{ANO_ATUAL}\{MES_ATUAL}\{EMPRESA}\{EMPRESA} ADIANTAMENTO.pdf
ORGANIZAR_BASE_DIR = r"W:\DOCUMENTOS ESCRITORIO\RH\AUTOMATIZADO\ADIANTAMENTO"

DELAY_BTN_S = 0.05
TYPE_INTERVAL_S = 0.02
TYPE_INTERVAL_CAMINHO_S = 0.10
WAIT_ANTES_DIGITAR_CAMINHO_S = 0.80
WAIT_DEPOIS_COLAR_CAMINHO_S = 0.40
WAIT_APOS_DIGITAR_S = 5.0


# =========================
# HELPERS
# =========================

def _sleep_btn() -> None:
    time.sleep(DELAY_BTN_S)


def _ensure_pag() -> None:
    global pag
    if pag is not None:
        return

    try:
        import pyautogui as _pag
    except Exception as exc:
        raise SystemExit("Dependencias faltando: pyautogui.") from exc

    pag = _pag


def _click_left(x: int, y: int, clicks: int = 1) -> None:
    for _ in range(max(1, int(clicks))):
        pag.click(int(x), int(y), button="left")
        _sleep_btn()


def _click_mouse2(x: int, y: int) -> None:
    # "mouse 2" interpretado como botao direito.
    pag.click(int(x), int(y), button="right")
    _sleep_btn()


def _press(key: str) -> None:
    pag.press(key)
    _sleep_btn()


def _hotkey(*keys: str) -> None:
    pag.hotkey(*keys)
    _sleep_btn()


def _type(text: str, interval_s: float = TYPE_INTERVAL_S) -> None:
    pag.typewrite(text, interval=float(interval_s))
    _sleep_btn()


def _ps_quote(text: str) -> str:
    # Aspas simples no PowerShell: para escapar, duplica.
    return "'" + text.replace("'", "''") + "'"


def _set_clipboard(text: str) -> None:
    # Evita dependencia extra (pyperclip). Usa PowerShell nativo.
    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            f"Set-Clipboard -Value {_ps_quote(text)}",
        ],
        check=True,
        capture_output=True,
        text=True,
    )


def _alt_m_e() -> None:
    pag.keyDown("alt")
    _sleep_btn()
    pag.press("m")
    _sleep_btn()
    pag.press("e")
    _sleep_btn()
    pag.keyUp("alt")
    _sleep_btn()


def _alt_p() -> None:
    pag.keyDown("alt")
    _sleep_btn()
    pag.press("p")
    _sleep_btn()
    pag.keyUp("alt")
    _sleep_btn()


def _nome_pasta_mes(dt: datetime) -> str:
    # Conforme solicitado: pasta do mes apenas com o numero (ex.: 02).
    return dt.strftime("%m")


def _limpar_nome_empresa(nome: str) -> str:
    """Normaliza para nome de pasta/arquivo.

    - remove acentos
    - remove caracteres invalidos do Windows
    - compacta espacos
    - deixa em MAIUSCULO
    """
    raw = (nome or "").strip()
    if not raw:
        return "EMPRESA"

    sem_acentos = "".join(
        ch
        for ch in unicodedata.normalize("NFKD", raw)
        if not unicodedata.combining(ch)
    )

    # Remove caracteres invalidos em nomes do Windows.
    sem_invalidos = re.sub(r'[<>:"/\\\\|?*]+', " ", sem_acentos)
    sem_invalidos = sem_invalidos.replace("\t", " ")
    sem_invalidos = re.sub(r"\s+", " ", sem_invalidos).strip()

    return (sem_invalidos or "EMPRESA").upper()


# =========================
# MODO ORGANIZAR
# =========================

def organizar_pastas(src_dir: str, dst_base_dir: str, mover: bool = False) -> int:
    """Modo apenas para formatar as pastas e organizar em:

    AUTOMATIZADO\\ADIANTAMENTO\\{ANO_ATUAL}\\{MES_ATUAL}\\

    Regras:
    - Pasta do mes: apenas numero (ex.: 02)
    - Dentro do mes: subpastas com nomes limpos de cada empresa
    - PDF renomeado para: NOME_DA_EMPRESA + ' ADIANTAMENTO'.pdf
    """
    dt = datetime.now()
    ano = dt.strftime("%Y")
    mes_pasta = _nome_pasta_mes(dt)

    if not os.path.isdir(src_dir):
        raise SystemExit(f"Pasta de origem nao existe: {src_dir}")

    dst_dir = os.path.join(dst_base_dir, ano, mes_pasta)
    os.makedirs(dst_dir, exist_ok=True)

    entries = sorted(os.listdir(src_dir))
    pdfs = [
        name
        for name in entries
        if os.path.isfile(os.path.join(src_dir, name)) and name.lower().endswith(".pdf")
    ]

    if not pdfs:
        print("Nenhum PDF para organizar.")
        print(f"Origem: {src_dir}")
        print(f"Destino: {dst_dir}")
        return 0

    processed = 0
    for name in pdfs:
        src = os.path.join(src_dir, name)
        base, _ext = os.path.splitext(name)

        empresa = _limpar_nome_empresa(base)
        pasta_empresa = os.path.join(dst_dir, empresa)
        os.makedirs(pasta_empresa, exist_ok=True)

        dst_name = f"{empresa} ADIANTAMENTO.pdf"
        dst = os.path.join(pasta_empresa, dst_name)

        if os.path.exists(dst):
            for n in range(1, 10000):
                candidate = os.path.join(
                    pasta_empresa, f"{empresa} ADIANTAMENTO__{n:04d}.pdf"
                )
                if not os.path.exists(candidate):
                    dst = candidate
                    break

        if mover:
            shutil.move(src, dst)
        else:
            shutil.copy2(src, dst)
        processed += 1

    acao = "movidos" if mover else "copiados"
    print(f"PDFs organizados ({acao}): {processed}")
    print(f"Origem: {src_dir}")
    print(f"Destino: {dst_dir}")
    return 0


def _ler_slk_mapa_empresas(slk_path: str) -> dict[str, str]:
    """Le um arquivo .slk (SYLK) simples e retorna um mapa:

    CODIGO (4 digitos, ex.: 0785) -> NOME (string)

    Observacao:
    - O arquivo usado aqui tem o codigo na coluna X1 e o nome na coluna X2.
    """
    if not os.path.isfile(slk_path):
        raise SystemExit(f"Arquivo SLK nao encontrado: {slk_path}")

    # SLK costuma ser ANSI/Latin-1 em exportacoes antigas; tenta UTF-8 e cai pra latin-1.
    try:
        with open(slk_path, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
    except UnicodeDecodeError:
        with open(slk_path, "r", encoding="latin-1") as f:
            lines = f.read().splitlines()

    # Parser basico para registros "C;...;K..."
    # Mantem estado de linha/coluna, pois o SLK pode omitir Y/X em algumas linhas.
    cur_row = None
    cur_col = None
    cells: dict[tuple[int, int], str] = {}

    for raw in lines:
        line = (raw or "").strip()
        if not line.startswith("C;"):
            continue

        parts = line.split(";")
        if len(parts) < 2:
            continue

        k_value = None
        # Procura o primeiro token iniciando com 'K' e junta o resto (se houver ';' no valor).
        for i, p in enumerate(parts):
            if p.startswith("Y") and p[1:].isdigit():
                cur_row = int(p[1:])
            elif p.startswith("X") and p[1:].isdigit():
                cur_col = int(p[1:])
            elif p.startswith("K"):
                k_value = ";".join(parts[i:])[1:]  # remove o 'K'
                break

        if cur_row is None or cur_col is None or k_value is None:
            continue

        v = k_value
        if v.startswith('"') and v.endswith('"') and len(v) >= 2:
            v = v[1:-1].replace('""', '"')
        v = (v or "").strip()
        cells[(cur_row, cur_col)] = v

    # Monta o mapa: col 1 = codigo; col 2 = nome.
    mapa: dict[str, str] = {}
    # Descobre o maior Y para iterar.
    if not cells:
        return mapa

    max_y = max(y for (y, _x) in cells.keys())
    for y in range(1, max_y + 1):
        cod = cells.get((y, 1), "").strip()
        nome = cells.get((y, 2), "").strip()
        if not cod or not nome:
            continue
        if not cod.isdigit():
            continue
        cod4 = cod.zfill(4)
        mapa[cod4] = nome

    return mapa


def organizar_pastas_por_codigo_slk(
    src_dir: str,
    dst_base_dir: str,
    slk_path: str,
    dry_run: bool = False,
    mover: bool = False,
) -> int:
    """Organiza PDFs gerados pelo fluxo 'salvar por empresa' do Adiantamento.

    Estrutura de origem observada:
    {SRC}\\{AAAAMM}\\...\\{CODIGO}\\...\\*.pdf

    Estrutura destino:
    {DST}\\{ANO}\\{MES}\\{EMPRESA}\\{EMPRESA} ADIANTAMENTO.pdf
    (se houver mais de um PDF, cria sufixo __0001, __0002, ...)
    """
    if not os.path.isdir(src_dir):
        raise SystemExit(f"Pasta de origem nao existe: {src_dir}")

    mapa = _ler_slk_mapa_empresas(slk_path)
    if not mapa:
        raise SystemExit(f"Nao foi possivel ler o mapa de empresas do SLK: {slk_path}")

    # Aceita src em 2 formatos:
    # - apontando para ...\\Arquivos (contendo pastas AAAAMM)
    # - apontando direto para ...\\Arquivos\\AAAAMM
    src_base = src_dir
    yyyymm_dirs = [
        d
        for d in os.listdir(src_base)
        if os.path.isdir(os.path.join(src_base, d)) and re.fullmatch(r"\d{6}", d)
    ]
    if re.fullmatch(r".*\\\d{6}$", src_base.replace("/", "\\")):
        yyyymm_dirs = [os.path.basename(src_base)]
        src_base = os.path.dirname(src_base)

    if not yyyymm_dirs:
        raise SystemExit(
            "Nao encontrei pasta AAAAMM dentro da origem. "
            f"Origem: {src_dir}"
        )

    processed = 0
    missing_map = []

    for yyyymm in sorted(yyyymm_dirs):
        ano = yyyymm[:4]
        mes_pasta = yyyymm[4:6]
        dst_dir = os.path.join(dst_base_dir, ano, mes_pasta)
        os.makedirs(dst_dir, exist_ok=True)

        month_dir = os.path.join(src_base, yyyymm)
        cod_paths: list[tuple[str, str]] = []
        for root, dirs, _files in os.walk(month_dir):
            for d in dirs:
                if re.fullmatch(r"\d{4}", d):
                    cod4 = d
                    cod_path = os.path.join(root, d)
                    cod_paths.append((cod4, cod_path))

        if not cod_paths:
            continue

        for cod4, cod_path in sorted(cod_paths, key=lambda x: (x[0], x[1])):
            pdfs = []
            for root, _dirs, files in os.walk(cod_path):
                for fn in files:
                    if fn.lower().endswith(".pdf"):
                        pdfs.append(os.path.join(root, fn))

            if not pdfs:
                # Ignora codigos sem PDF (comportamento solicitado).
                continue

            nome_raw = mapa.get(cod4, "").strip()
            if not nome_raw:
                missing_map.append(cod4)
                nome_raw = cod4

            empresa = _limpar_nome_empresa(nome_raw)
            pasta_empresa = os.path.join(dst_dir, empresa)
            os.makedirs(pasta_empresa, exist_ok=True)

            for src_pdf in sorted(pdfs):
                dst_name = f"{empresa} ADIANTAMENTO.pdf"
                dst = os.path.join(pasta_empresa, dst_name)

                if os.path.exists(dst):
                    for n in range(1, 10000):
                        candidate = os.path.join(
                            pasta_empresa, f"{empresa} ADIANTAMENTO__{n:04d}.pdf"
                        )
                        if not os.path.exists(candidate):
                            dst = candidate
                            break

                if dry_run:
                    print(f"[DRY-RUN] {src_pdf} -> {dst}")
                else:
                    if mover:
                        shutil.move(src_pdf, dst)
                    else:
                        shutil.copy2(src_pdf, dst)
                processed += 1

    acao = "movidos" if mover else "copiados"
    print(f"PDFs organizados ({acao}): {processed}")
    print(f"Origem: {src_dir}")
    print(f"Destino: {dst_base_dir}")
    print(f"SLK: {slk_path}")

    if missing_map:
        uniq = sorted(set(missing_map))
        print("Aviso: codigos nao encontrados no SLK (usando o proprio codigo como nome):")
        for cod4 in uniq:
            print(f" - {cod4}")

    return 0


def gerar_exemplo_organizador() -> int:
    """Gera um exemplo para testar o organizador."""
    base_dir = os.path.dirname(__file__)

    exemplo_src = os.path.join(base_dir, "Arquivos_EXEMPLO")
    exemplo_dst_base = os.path.join(base_dir, "AUTOMATIZADO_EXEMPLO", "ADIANTAMENTO")

    os.makedirs(exemplo_src, exist_ok=True)
    os.makedirs(exemplo_dst_base, exist_ok=True)

    samples = [
        "ACME & Filhos LTDA.pdf",
        "Sao Joao Comercio (Matriz).PDF",
        "Otica Uniao - Filial 01.pdf",
    ]

    # Arquivo minimo para fins de teste de movimentacao/renome.
    pdf_stub = b"%PDF-1.4\n% stub\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n"

    for fname in samples:
        path = os.path.join(exemplo_src, fname)
        if not os.path.exists(path):
            with open(path, "wb") as f:
                f.write(pdf_stub)

    print("Exemplo criado.")
    print(f"Origem exemplo: {exemplo_src}")
    print(f"Destino base exemplo: {exemplo_dst_base}")
    print("Para testar a organizacao no exemplo, rode:")
    print(
        f'python "{os.path.join(base_dir, "Adiantamento.py")}" --organizar --src "{exemplo_src}" --dst "{exemplo_dst_base}"'
    )
    return 0


# =========================
# MODO AUTOMACAO
# =========================

def rodar_automacao() -> int:
    _ensure_pag()
    pag.FAILSAFE = True

    # 1) Foca na tela completamente.
    _click_left(FOCUS_X, FOCUS_Y, clicks=FOCUS_CLICKS)

    # 2) ALT M E
    _alt_m_e()

    # 3) Sem clicar em nada: digitar 1, MM/YY (atual), 1, 30
    mes_ano_mm_yy = datetime.now().strftime("%m/%y")

    _type("1")
    _press("enter")

    _type(mes_ano_mm_yy)
    _press("enter")

    _type("1")
    _press("enter")

    _type("30")
    _press("enter")

    # 4) Espera 5 segundos
    time.sleep(WAIT_APOS_DIGITAR_S)

    # 5) Fluxo com mouse 2 em coordenadas + mouse 2 5px a esquerda + teclas
    for x, y, keys in MOUSE2_SEQS:
        _click_mouse2(x, y)
        _click_mouse2(x - MOUSE2_LEFT_SHIFT_PX, y)
        for k in keys:
            _press(k)

    # 6) Cliques finais
    _click_left(*CLICK_PARAMETRO)
    _click_left(*CLICK_RELATORIOS_MENSAIS)

    # Antes de marcar "salvar arquivos por empresa": clicar, 2 backspaces e digitar 1.
    _click_left(*CLICK_CAMPO_ANTES_SALVAR_POR_EMPRESA)
    _press("backspace")
    _press("backspace")
    _type("1")

    _click_left(*CLICK_SALVAR_ARQ_POR_EMPRESA)

    _click_left(*CLICK_CAMINHO_DIRETORIO)
    time.sleep(WAIT_ANTES_DIGITAR_CAMINHO_S)
    _hotkey("ctrl", "a")
    try:
        # Colar e mais confiavel que digitar em campos que perdem foco.
        _set_clipboard(DIRETORIO_ARQUIVOS)
        _hotkey("ctrl", "v")
        time.sleep(WAIT_DEPOIS_COLAR_CAMINHO_S)
    except Exception:
        # Fallback: digitar bem devagar.
        _type(DIRETORIO_ARQUIVOS, interval_s=TYPE_INTERVAL_CAMINHO_S)

    # 7) ALT P
    _alt_p()

    return 0


# =========================
# MAIN
# =========================

def _get_arg_value(flag: str) -> str | None:
    flag_l = flag.lower()
    argv = [a for a in sys.argv[1:]]
    for i, a in enumerate(argv):
        if a.lower() == flag_l and i + 1 < len(argv):
            return argv[i + 1]
    return None


def main() -> int:
    # Modos:
    # - default: automacao UI
    # - --organizar: somente organizar pastas/arquivos
    #   flags extras: --src "..."  --dst "..."
    # - --gerar-exemplo: cria Arquivos_EXEMPLO e mostra comando de teste

    argv_l = [a.lower() for a in sys.argv[1:]]

    if "--gerar-exemplo" in argv_l or "/gerar-exemplo" in argv_l:
        return gerar_exemplo_organizador()

    if "--organizar" in argv_l or "/organizar" in argv_l:
        src = _get_arg_value("--src") or _get_arg_value("/src") or DIRETORIO_ARQUIVOS
        dst = _get_arg_value("--dst") or _get_arg_value("/dst") or ORGANIZAR_BASE_DIR
        slk = _get_arg_value("--slk") or _get_arg_value("/slk")
        dry_run = ("--dry-run" in argv_l) or ("/dry-run" in argv_l)
        mover = ("--mover" in argv_l) or ("/mover" in argv_l)
        if slk:
            return organizar_pastas_por_codigo_slk(
                src_dir=src,
                dst_base_dir=dst,
                slk_path=slk,
                dry_run=dry_run,
                mover=mover,
            )
        return organizar_pastas(src_dir=src, dst_base_dir=dst, mover=mover)

    return rodar_automacao()


if __name__ == "__main__":
    raise SystemExit(main())
