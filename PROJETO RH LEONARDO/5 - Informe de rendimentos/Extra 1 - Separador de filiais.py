import argparse
import os
import csv
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

try:
    from pypdf import PdfReader
except Exception as exc:
    raise SystemExit("pypdf is required to read the PDF.") from exc


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CODES_CSV = (
    r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\5 - Informe de rendimentos\lista bruta empresas.csv"
)
DEFAULT_PDF = (
    r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\PROJETO RH LEONARDO\5 - Informe de rendimentos\RELAÇÃO DE EMPRESAS.pdf"
)
DEFAULT_OUTPUT = BASE_DIR / "empresas.csv"


STATUS_SUFFIX_RE = re.compile(r"(Ativo|Inativo|Suspenso|Baixado)\s*$", re.IGNORECASE)


@dataclass
class Company:
    code: str
    name: str
    estabs: Set[str] = field(default_factory=set)


def normalize_code(value: str) -> Optional[str]:
    if value is None:
        return None
    m = re.match(r"\s*(\d+)", str(value))
    if not m:
        return None
    return str(int(m.group(1)))


def _sniff_delimiter(sample: str) -> str:
    try:
        return csv.Sniffer().sniff(sample, delimiters=";,").delimiter
    except Exception:
        return ";"


def load_codes_list(path: Path) -> List[Tuple[str, str]]:
    if not path.exists():
        raise SystemExit(f"Arquivo nao encontrado: {path}")
    try:
        text = path.read_text(encoding="utf-8-sig")
        file_encoding = "utf-8-sig"
    except UnicodeDecodeError:
        text = path.read_text(encoding="latin-1", errors="ignore")
        file_encoding = "latin-1"
    delimiter = _sniff_delimiter(text[:2048])
    rows = []
    with path.open("r", encoding=file_encoding, errors="ignore", newline="") as f:
        reader = csv.reader(f, delimiter=delimiter)
        for row in reader:
            if not row:
                continue
            rows.append([c.strip() for c in row])
    if not rows:
        return []

    header = [c.lower() for c in rows[0]]
    has_header = any("codigo" in c for c in header) or any("respons" in c for c in header)
    if has_header:
        rows = rows[1:]

    codes: List[Tuple[str, str]] = []
    seen: Set[str] = set()
    for row in rows:
        if not row:
            continue
        code_raw = row[0] if len(row) >= 1 else ""
        resp = row[1] if len(row) >= 2 else ""
        code = normalize_code(code_raw)
        if not code:
            continue
        if code in seen:
            continue
        seen.add(code)
        codes.append((code, resp.strip()))
    return codes


def _extract_status_suffix(text: str) -> Tuple[str, str]:
    m = STATUS_SUFFIX_RE.search(text or "")
    if not m:
        return "", text.strip()
    status = m.group(1)
    rest = (text[: m.start()] or "").strip()
    return status, rest


def _find_doc_index(tokens: List[str]) -> int:
    for i, tok in enumerate(tokens):
        digits = re.sub(r"\D", "", tok or "")
        if len(digits) >= 8:
            return i
    return -1


def _extract_estab_from_doc(doc_token: str) -> Optional[str]:
    digits = re.sub(r"\D", "", doc_token or "")
    if len(digits) < 12:
        return None
    if len(digits) >= 14:
        digits = digits[-14:]
        estab = digits[8:12]
        try:
            return str(int(estab))
        except Exception:
            return estab
    return None


def _norm_name(value: str) -> str:
    text = (value or "").strip().lower()
    text = re.sub(r"[\\W_]+", " ", text, flags=re.UNICODE)
    return " ".join(text.split())


def _names_match(current_name: str, line_name: str) -> bool:
    if not current_name or not line_name:
        return False
    a = _norm_name(current_name)
    b = _norm_name(line_name)
    if not a or not b:
        return False
    if a in b or b in a:
        return True
    # tolera sufixos como "filial"
    b_clean = b.replace(" filial", "")
    return a in b_clean or b_clean in a


def parse_pdf(pdf_path: Path) -> Dict[str, Company]:
    if not pdf_path.exists():
        raise SystemExit(f"PDF nao encontrado: {pdf_path}")
    reader = PdfReader(str(pdf_path))
    companies: Dict[str, Company] = {}
    current_code: Optional[str] = None

    mismatches: list[tuple[str, str, str]] = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for raw in text.splitlines():
            line = raw.strip()
            if not line:
                continue
            m = re.match(r"^(?P<code>\d{4,5})\s+(?P<rest>.+)$", line)
            if not m:
                continue
            code_raw = m.group("code")
            rest_full = m.group("rest").strip()
            code = normalize_code(code_raw)
            if not code:
                continue

            status, rest = _extract_status_suffix(rest_full)
            tokens = rest.split()

            doc_idx = _find_doc_index(tokens)

            if doc_idx == -1:
                # Sem documento: assume cabecalho da empresa
                name = rest.strip()
                current_code = code
                if code not in companies:
                    companies[code] = Company(code=code, name=name)
                else:
                    if not companies[code].name and name:
                        companies[code].name = name
                continue

            name = " ".join(tokens[:doc_idx]).strip()

            # Se o documento for o ultimo token, trata como cabecalho
            if doc_idx == len(tokens) - 1:
                current_code = code
                if code not in companies:
                    companies[code] = Company(code=code, name=name)
                else:
                    if not companies[code].name and name:
                        companies[code].name = name
                continue

            # Caso contrario, tenta validar se pertence ao cabecalho atual
            if current_code and _names_match(companies.get(current_code, Company("", "")).name, name):
                companies.setdefault(current_code, Company(code=current_code, name=""))
                status_l = status.lower() if status else ""
                if status_l not in ("inativo", "baixado", "suspenso"):
                    companies[current_code].estabs.add(str(int(code_raw)))
                continue

            # Se o nome nao bate, trata como novo cabecalho (evita "filiais" falsas)
            if current_code:
                mismatches.append((current_code, companies.get(current_code, Company("", "")).name, name))
            current_code = code
            if code not in companies:
                companies[code] = Company(code=code, name=name)
            else:
                if not companies[code].name and name:
                    companies[code].name = name

    # salva relatorio de inconsistencias
    if mismatches:
        report = pdf_path.with_name("relatorio_mismatches_filiais.txt")
        lines = ["codigo_atual;nome_atual;nome_linha"]
        for code, cur_name, line_name in mismatches:
            lines.append(f"{code};{cur_name};{line_name}")
        report.write_text("\n".join(lines), encoding="utf-8")

    return companies


def write_empresas_csv(
    output: Path, codes: List[Tuple[str, str]], companies: Dict[str, Company]
) -> None:
    with output.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["CÓDIGO", "NOME", "ESTABELECIMENTO", "FUNCIONARIO"])

        responsavel_order: Dict[str, int] = {}
        for _, resp in codes:
            key = resp.strip()
            if key not in responsavel_order:
                responsavel_order[key] = len(responsavel_order)

        def _sort_key(item: Tuple[str, str]) -> Tuple[int, int]:
            code, resp = item
            resp_key = resp.strip()
            resp_idx = responsavel_order.get(resp_key, 10_000)
            return (resp_idx, int(code) if code.isdigit() else 0)

        for code, resp in sorted(codes, key=_sort_key):
            data = companies.get(code)
            if not data:
                writer.writerow([code, "", "1", resp])
                print(f"[warn] codigo nao encontrado no PDF: {code}")
                continue
            name = data.name or code
            estabs = sorted(
                {str(int(e)) for e in data.estabs if str(e).isdigit()},
                key=lambda x: int(x),
            )
            estab_str = ",".join(estabs) if estabs else "1"
            writer.writerow([code, name, estab_str, resp])


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--codes-csv",
        default=DEFAULT_CODES_CSV,
        help="Arquivo com a lista bruta de codigos (um por linha).",
    )
    ap.add_argument(
        "--pdf",
        default=DEFAULT_PDF,
        help="PDF RELACAO DE EMPRESAS.",
    )
    ap.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT),
        help="CSV de saida (empresas.csv).",
    )
    args = ap.parse_args()

    codes = load_codes_list(Path(args.codes_csv))
    companies = parse_pdf(Path(args.pdf))
    output_path = Path(args.output)
    write_empresas_csv(output_path, codes, companies)
    try:
        os.startfile(output_path)
    except Exception:
        pass
    print(f"[ok] salvo: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
