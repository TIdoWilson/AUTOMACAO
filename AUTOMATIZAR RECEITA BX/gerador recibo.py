import re
import time
import argparse
from pathlib import Path
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

# ======================================================
# Regras de nome (entrada)
#   PISCOFINS_"data inicial"_"data final"_"CNPJ"_(opcional CNPJ da SCP)_"Original/Retificadora"
#   Ex.: PISCOFINS_20260101_20260131_60087437000136_63478445000100_Original
#
# Regras de nome (saída)
#   "RECIBO CNPJ [CNPJ_SCP SCP] MES/ANO(mm/aaaa) SPED CONTRIBUICOES ORIGINAL/RETIFICADORA.pdf"
#   - O bloco "CNPJ_SCP SCP" só aparece se houver 2º CNPJ no nome de entrada.
#   - ORIGINAL/RETIFICADORA deve refletir o que está no nome de entrada.
#   - Mês/ano é derivado da "data final" (YYYYMMDD) do nome de entrada.
# ======================================================

# ---------------- parsing do nome do arquivo ----------------

# CNPJ: 14 dígitos, isolado (evita pegar trechos de datas/seqs maiores)
CNPJ_RE = re.compile(r"(?<!\d)(\d{14})(?!\d)")

# Datas no padrão YYYYMMDD, isoladas
DATE_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")

def _tokens_nome(stem: str) -> List[str]:
    # Separação segura do nome (sem extensão)
    return stem.split("_")

def _indice_tipo(tokens: List[str]) -> Optional[int]:
    for i, t in enumerate(tokens):
        tl = t.strip().lower()
        if tl in ("original", "retificadora"):
            return i
    return None

def extrair_cnpjs(stem: str) -> List[str]:
    """Retorna os CNPJs do *bloco principal* do nome (antes de ORIGINAL/RETIFICADORA).

    Importante: alguns arquivos incluem sufixos após ORIGINAL/RETIFICADORA (timestamp/hash etc.)
    que podem conter sequências de 14 dígitos — isso NÃO é SCP.
    """
    tokens = _tokens_nome(stem)
    idx_tipo = _indice_tipo(tokens)

    # Preferência: CNPJs como tokens antes do tipo
    if idx_tipo is not None:
        candidatos = []
        for t in tokens[:idx_tipo]:
            if re.fullmatch(r"\d{14}", t):
                candidatos.append(t)
        if candidatos:
            return candidatos

    # Fallback: regex global (pode capturar ruído, mas evita quebrar caso formato esteja diferente)
    return CNPJ_RE.findall(stem)

def extrair_periodo_mm_aaaa(stem: str) -> str:
    """Extrai MM/AAAA a partir da *data final* do nome.

    Regra: usa o 3º token (data final) no padrão:
      PISCOFINS_YYYYMMDD_YYYYMMDD_...
    Isso evita confundir com datas de timestamp após ORIGINAL/RETIFICADORA.
    """
    tokens = _tokens_nome(stem)
    idx_tipo = _indice_tipo(tokens)
    bloco_principal = tokens[:idx_tipo] if idx_tipo is not None else tokens

    # Regra principal: "data final" no 3o token (aceita varios formatos).
    if len(bloco_principal) >= 3:
        periodo = _token_para_mm_aaaa(bloco_principal[2], aceitar_data=True)
        if periodo:
            return periodo

    # Fallback: procura do fim para o inicio no bloco principal.
    for t in reversed(bloco_principal):
        periodo = _token_para_mm_aaaa(t, aceitar_data=True)
        if periodo:
            return periodo

    # Fallback legado: datas de 8 digitos no nome inteiro.
    datas = DATE_RE.findall(stem)
    if len(datas) >= 2:
        periodo = _token_para_mm_aaaa(datas[1], aceitar_data=True)
        if periodo:
            return periodo

    return "000000"


def _token_para_mm_aaaa(token: str, aceitar_data: bool = True) -> Optional[str]:
    """Converte um token para MMYYYY, aceitando mes/ano ou data completa."""
    t = token.strip()
    if not t:
        return None

    # YYYY-MM-DD / YYYY/MM/DD
    m = re.fullmatch(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", t)
    if m and aceitar_data:
        ano, mes, dia = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if _data_valida(ano, mes, dia):
            return f"{mes:02d}{ano:04d}"

    # DD-MM-YYYY / DD/MM/YYYY
    m = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", t)
    if m and aceitar_data:
        dia, mes, ano = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if _data_valida(ano, mes, dia):
            return f"{mes:02d}{ano:04d}"

    # MM-YYYY / MM/YYYY
    m = re.fullmatch(r"(\d{1,2})[/-](\d{4})", t)
    if m:
        mes, ano = int(m.group(1)), int(m.group(2))
        if _mes_ano_valido(mes, ano):
            return f"{mes:02d}{ano:04d}"

    # MM-YY / MM/YY
    m = re.fullmatch(r"(\d{1,2})[/-](\d{2})", t)
    if m:
        mes, ano = int(m.group(1)), 2000 + int(m.group(2))
        if _mes_ano_valido(mes, ano):
            return f"{mes:02d}{ano:04d}"

    # Apenas digitos
    d = re.sub(r"\D", "", t)
    if not d:
        return None

    # YYYYMMDD
    if len(d) == 8 and aceitar_data:
        ano, mes, dia = int(d[0:4]), int(d[4:6]), int(d[6:8])
        if _data_valida(ano, mes, dia):
            return f"{mes:02d}{ano:04d}"

    # DDMMYYYY
    if len(d) == 8 and aceitar_data:
        dia, mes, ano = int(d[0:2]), int(d[2:4]), int(d[4:8])
        if _data_valida(ano, mes, dia):
            return f"{mes:02d}{ano:04d}"

    # MMYYYY
    if len(d) == 6:
        mes, ano = int(d[0:2]), int(d[2:6])
        if _mes_ano_valido(mes, ano):
            return f"{mes:02d}{ano:04d}"

    # MMYY
    if len(d) == 4:
        mes, ano = int(d[0:2]), 2000 + int(d[2:4])
        if _mes_ano_valido(mes, ano):
            return f"{mes:02d}{ano:04d}"

    return None


def _mes_ano_valido(mes: int, ano: int) -> bool:
    return 1 <= mes <= 12 and 1900 <= ano <= 2100


def _data_valida(ano: int, mes: int, dia: int) -> bool:
    try:
        datetime(ano, mes, dia)
        return True
    except ValueError:
        return False

def extrair_tipo_original_retificadora(stem: str) -> str:
    """Retorna 'ORIGINAL' ou 'RETIFICADORA' conforme token do nome."""
    tokens = _tokens_nome(stem)
    idx_tipo = _indice_tipo(tokens)
    if idx_tipo is not None:
        tl = tokens[idx_tipo].strip().lower()
        if tl == "retificadora":
            return "RETIFICADORA"
        if tl == "original":
            return "ORIGINAL"

    # fallback
    s = stem.lower()
    if "retificadora" in s:
        return "RETIFICADORA"
    return "ORIGINAL"
    # fallback seguro
    return "ORIGINAL"

def garantir_unico(path: Path) -> Path:
    """Se o arquivo já existir, adiciona sufixo incremental: (2), (3), ..."""
    if not path.exists():
        return path
    base = path.with_suffix("")
    ext = path.suffix
    i = 2
    while True:
        candidato = Path(f"{base} ({i}){ext}")
        if not candidato.exists():
            return candidato
        i += 1

def montar_nome_saida(stem: str) -> Tuple[str, str, Optional[str], str]:
    """
    Retorna:
      - nome_base (sem .pdf)
      - cnpj_principal
      - cnpj_scp (ou None)
      - tipo (ORIGINAL/RETIFICADORA)
    """
    cnpjs = extrair_cnpjs(stem)
    cnpj_principal = cnpjs[0] if len(cnpjs) >= 1 else "SEM_CNPJ"
    cnpj_scp = cnpjs[1] if len(cnpjs) >= 2 else None

    mm_aaaa = extrair_periodo_mm_aaaa(stem)
    tipo = extrair_tipo_original_retificadora(stem)

    partes = ["RECIBO", cnpj_principal]
    if cnpj_scp:
        partes.extend([cnpj_scp, "SCP"])
    partes.extend([mm_aaaa, "SPED CONTRIBUICOES", tipo])

    nome_base = " ".join(partes)
    return nome_base, cnpj_principal, cnpj_scp, tipo

# ---------------- Seleção de pasta (GUI) ----------------
def selecionar_pasta() -> Optional[Path]:
    from tkinter import Tk, filedialog

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    pasta = filedialog.askdirectory(
        title="Selecione a pasta com os arquivos .txt e .rec"
    )
    root.destroy()

    if not pasta:
        return None
    return Path(pasta)

# ---------------- Montagem de duplas .txt + .rec ----------------
@dataclass(frozen=True)
class ParSped:
    txt: Path
    rec: Path
    out: Path
    cnpj: str
    cnpj_scp: Optional[str]
    periodo: str
    tipo: str

def encontrar_pares(root: Path, recursive: bool = True) -> Tuple[List[ParSped], List[Path]]:
    pattern = "**/*" if recursive else "*"
    candidatos = [
        p for p in root.glob(pattern)
        if p.is_file() and p.suffix.lower() in {".txt", ".rec"}
    ]

    por_stem: Dict[Tuple[Path, str], Dict[str, Path]] = {}
    for p in candidatos:
        key = (p.parent, p.stem)
        por_stem.setdefault(key, {})
        por_stem[key][p.suffix.lower()] = p

    pares: List[ParSped] = []
    sem_par: List[Path] = []

    for (folder, stem), exts in sorted(
        por_stem.items(),
        key=lambda x: (str(x[0][0]).lower(), x[0][1].lower())
    ):
        txt = exts.get(".txt")
        rec = exts.get(".rec")

        if txt and rec:
            nome_base, cnpj, cnpj_scp, tipo = montar_nome_saida(stem)
            out_path = garantir_unico(folder / f"{nome_base}.pdf")

            # período para log/diagnóstico
            periodo = extrair_periodo_mm_aaaa(stem)

            pares.append(
                ParSped(
                    txt=txt,
                    rec=rec,
                    out=out_path,
                    cnpj=cnpj,
                    cnpj_scp=cnpj_scp,
                    periodo=periodo,
                    tipo=tipo,
                )
            )
        else:
            # guarda somente o que está faltando par (diagnóstico)
            if txt:
                sem_par.append(txt)
            if rec:
                sem_par.append(rec)

    return pares, sem_par


# ---------------- UI Automation (Windows) ----------------
def _escape_send_keys_text(text: str) -> str:
    """Escapa texto para pywinauto.keyboard.send_keys (digitação literal).

    send_keys interpreta alguns caracteres como meta (ex.: + ^ % ~ e chaves).
    Esta função força a digitação literal desses caracteres.
    """
    # chaves
    text = text.replace("{", "{{}").replace("}", "{}}")
    # metacaracteres
    text = (
        text.replace("+", "{+}")
            .replace("^", "{^}")
            .replace("%", "{%}")
            .replace("~", "{~}")
    )
    return text

def gerar_recibo_ui(

    txt_path: Path,
    rec_path: Path,
    out_path: Path,
    titulo_janela: str = "Sistema Público de Escrituração Digital - EFD Contribuições",
    delay: float = 0.25,
    bigdelay: float = 1.5,
    type_pause: float = 0.001,
    type_with_pause: bool = True,
) -> None:
    from pywinauto import Desktop
    from pywinauto.keyboard import send_keys

    w = Desktop(backend="win32").window(title=titulo_janela)
    if not w.exists(timeout=5):
        raise RuntimeError(
            f"Janela não encontrada: '{titulo_janela}'. Abra o EFD Contribuições antes de rodar."
        )
    w.set_focus()
    time.sleep(delay)

    # Abre rotina de recibo
    send_keys("^r")
    time.sleep(delay)
    send_keys("{ENTER}")
    time.sleep(delay)
    send_keys(
        _escape_send_keys_text(str(txt_path)),
        with_spaces=True,
        pause=(type_pause if type_with_pause else 0),
    )
    time.sleep(delay)
    send_keys("{ENTER}")
    
    time.sleep(delay)
    send_keys("{SPACE}")
    time.sleep(delay)
    
    time.sleep(bigdelay) 
       
    # REC (digitação rápida)
    send_keys("{ENTER}")
   
    send_keys(
        _escape_send_keys_text(str(rec_path)),
        with_spaces=True,
        pause=(type_pause if type_with_pause else 0),
    )

    time.sleep(delay)
    send_keys("{ENTER}")
    time.sleep(delay)

    send_keys("{SPACE}")
    time.sleep(2)
    send_keys("{SPACE}")
    time.sleep(delay)

    # Saída PDF (digitação rápida)
    send_keys(
        _escape_send_keys_text(str(out_path)),
        with_spaces=True,
        pause=(type_pause if type_with_pause else 0),
    )
    time.sleep(delay)
    send_keys("{ENTER}")
    time.sleep(3)

    send_keys("{ESC}")
    time.sleep(delay)

# ---------------- Main ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="Só listar pares (não automatiza UI)")
    ap.add_argument("--no-recursive", action="store_true", help="Não varrer subpastas")
    ap.add_argument("--delay", type=float, default=0.10, help="Delay entre passos (segundos)")
    ap.add_argument(
        "--type-pause",
        type=float,
        default=0.001,
        help="Pausa entre caracteres no send_keys (segundos). Quanto menor, mais rápido.",
    )
    ap.add_argument(
        "--type-no-pause",
        action="store_true",
        help="Desativa pausa entre caracteres (mais rápido, mas pode falhar em algumas máquinas).",
    )
    args = ap.parse_args()

    root = selecionar_pasta()
    if root is None:
        print("Seleção cancelada. Encerrando.")
        return

    if not root.exists():
        raise SystemExit(f"Pasta não existe: {root}")

    pares, sem_par = encontrar_pares(root, recursive=not args.no_recursive)

    print(f"\nPasta selecionada: {root}")
    print(f"Pares encontrados: {len(pares)}")
    for i, p in enumerate(pares, 1):
        scp_info = f" | SCP={p.cnpj_scp}" if p.cnpj_scp else ""
        print(f"{i:03d} | CNPJ={p.cnpj}{scp_info} | PERIODO={p.periodo} | TIPO={p.tipo}")
        print(f"      TXT={p.txt.name} | REC={p.rec.name}")
        print(f"      OUT={p.out}")

    if sem_par:
        print(f"\nArquivos sem par: {len(sem_par)}")
        for p in sem_par:
            print(f" - {p}")

    if args.dry_run:
        return

    print("\nIniciando geração via UI... (deixe o EFD Contribuições aberto e não use teclado/mouse)")
    for i, p in enumerate(pares, 1):
        print(f"[{i}/{len(pares)}] Gerando: {p.out.name}")
        gerar_recibo_ui(
            p.txt,
            p.rec,
            p.out,
            delay=args.delay,
            type_pause=args.type_pause,
            type_with_pause=(not args.type_no_pause),
        )

    print("\nConcluído.")

if __name__ == "__main__":
    main()
