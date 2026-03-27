import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path, PureWindowsPath


SCRIPT_DIR = Path(__file__).resolve().parent
BACKUP_BASE_PADRAO = Path(r"C:\Users\User\OneDrive\Área de Trabalho\Backups RH")
MAX_ARQUIVOS_DIAGNOSTICO = 15


def _log(log_fh, msg: str) -> None:
    log_fh.write(msg)
    log_fh.flush()


def localizar_ghostscript(gs_cli: str | None) -> str:
    if gs_cli:
        p = Path(gs_cli)
        if p.exists() and p.is_file():
            return str(p)
        if p.exists() and p.is_dir():
            candidatos_local = [
                p / "bin" / "gswin64c.exe",
                p / "bin" / "gswin64.exe",
                p / "gswin64c.exe",
                p / "gswin64.exe",
                p / "bin" / "gspcl6win64.exe",
                p / "gspcl6win64.exe",
            ]
            for c in candidatos_local:
                if c.exists():
                    return str(c)
        raise FileNotFoundError(f"Ghostscript nao encontrado em: {gs_cli}")

    candidatos = [
        Path(r"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.03.1\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.07.0\bin\gswin64c.exe"),
    ]
    for c in candidatos:
        if c.exists():
            return str(c)

    return "gswin64c"


def localizar_ocrmypdf(ocr_cli: str | None) -> str:
    if ocr_cli:
        p = Path(ocr_cli)
        if p.exists() and p.is_file():
            return str(p)
        if p.exists() and p.is_dir():
            for cand in [
                p / "ocrmypdf.exe",
                p / "Scripts" / "ocrmypdf.exe",
            ]:
                if cand.exists():
                    return str(cand)
        raise FileNotFoundError(f"OCRmyPDF nao encontrado em: {ocr_cli}")
    return "ocrmypdf"


def extrair_texto_len(pdf_path: Path) -> int:
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception:
        raise RuntimeError(
            "Dependencia ausente: pypdf. Instale com: pip install pypdf"
        )

    try:
        reader = PdfReader(str(pdf_path))
        total = 0
        for page in reader.pages:
            txt = page.extract_text() or ""
            total += len(txt.strip())
        return total
    except Exception:
        return -1


def _normalizar_pdf_com_pypdf(src_pdf: Path, dst_pdf: Path) -> tuple[bool, str]:
    try:
        from pypdf import PdfReader, PdfWriter  # type: ignore
    except Exception as exc:
        return False, f"pypdf indisponivel: {exc!r}"

    try:
        reader = PdfReader(str(src_pdf))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        with open(dst_pdf, "wb") as f:
            writer.write(f)
        return True, ""
    except Exception as exc:
        return False, f"falha ao normalizar com pypdf: {exc!r}"


def _gs_versao(gs_exe: str) -> str:
    try:
        p = _run_subprocess([gs_exe, "-v"], timeout=15)
        txt = ((p.stdout or "") + "\n" + (p.stderr or "")).strip()
        return txt or f"(sem saida, exit={p.returncode})"
    except Exception as exc:
        return f"(nao foi possivel obter versao: {exc!r})"


def _run_subprocess(cmd: list[str], timeout: int | None = None) -> subprocess.CompletedProcess:
    kwargs = {
        "capture_output": True,
        "text": True,
    }
    if timeout is not None:
        kwargs["timeout"] = timeout
    if os.name == "nt":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        kwargs["startupinfo"] = si
    return subprocess.run(cmd, **kwargs)


def _executaveis_gs_para_tentar(gs_exe: str) -> list[str]:
    base = Path(gs_exe)
    cands = [base]
    extras = [
        base.with_name("gspcl6win64.exe"),
        base.parent / "gspcl6win64.exe",
    ]
    for e in extras:
        if e.exists():
            cands.append(e)
    # remove duplicados preservando ordem
    out: list[str] = []
    seen: set[str] = set()
    for c in cands:
        s = str(c)
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out


def converter_pdfa_ghostscript(gs_exe: str, src_pdf: Path, out_pdf: Path) -> tuple[bool, str]:
    # Perfis em ordem de preferencia para maximizar compatibilidade com GS 10.07+
    perfis = [
        # Perfil principal PDF/A-2 (sem -dNEWPDF, removido no GS novo)
        [
            "-dPDFA=2",
            "-dBATCH",
            "-dNOPAUSE",
            "-dNOPROMPT",
            "-sDEVICE=pdfwrite",
            "-sProcessColorModel=DeviceRGB",
            "-sColorConversionStrategy=RGB",
            "-sPDFACompatibilityPolicy=1",
            f"-sOutputFile={out_pdf}",
            "-f",
            str(src_pdf),
        ],
        # Fallback PDF/A-1
        [
            "-dPDFA=1",
            "-dBATCH",
            "-dNOPAUSE",
            "-dNOPROMPT",
            "-sDEVICE=pdfwrite",
            "-sProcessColorModel=DeviceRGB",
            "-sColorConversionStrategy=RGB",
            "-sPDFACompatibilityPolicy=1",
            f"-sOutputFile={out_pdf}",
            "-f",
            str(src_pdf),
        ],
        # Fallback mantendo cores originais (alguns PDFs quebram na estrategia RGB)
        [
            "-dPDFA=2",
            "-dBATCH",
            "-dNOPAUSE",
            "-dNOPROMPT",
            "-sDEVICE=pdfwrite",
            "-sColorConversionStrategy=LeaveColorUnchanged",
            "-sPDFACompatibilityPolicy=1",
            f"-sOutputFile={out_pdf}",
            "-f",
            str(src_pdf),
        ],
        # Ultimo fallback PDF/A-1 + cores originais
        [
            "-dPDFA=1",
            "-dBATCH",
            "-dNOPAUSE",
            "-dNOPROMPT",
            "-sDEVICE=pdfwrite",
            "-sColorConversionStrategy=LeaveColorUnchanged",
            "-sPDFACompatibilityPolicy=1",
            f"-sOutputFile={out_pdf}",
            "-f",
            str(src_pdf),
        ],
    ]

    erros: list[str] = []
    executaveis = _executaveis_gs_para_tentar(gs_exe)
    t = 0
    for exe in executaveis:
        for perfil in perfis:
            t += 1
            cmd = [exe] + perfil
            proc = _run_subprocess(cmd)
            if proc.returncode == 0 and out_pdf.exists():
                return True, ""
            err = (
                f"tentativa {t}\n"
                f"cmd: {' '.join(cmd)}\n"
                f"exit_code: {proc.returncode}\n"
                f"--- STDERR ---\n{(proc.stderr or '').strip()}\n"
                f"--- STDOUT ---\n{(proc.stdout or '').strip()}\n"
                f"--------------"
            )
            erros.append(err)
            out_pdf.unlink(missing_ok=True)

    return False, "\n\n".join(erros)


def converter_pdfa_ocrmypdf(ocr_exe: str, src_pdf: Path, out_pdf: Path) -> tuple[bool, str]:
    # Sem OCR forcado para preservar texto existente; --skip-text evita rasterizar texto ja presente.
    cmd = [
        ocr_exe,
        "--skip-text",
        "--output-type",
        "pdfa-2",
        "--optimize",
        "0",
        str(src_pdf),
        str(out_pdf),
    ]
    proc = _run_subprocess(cmd)
    if proc.returncode == 0 and out_pdf.exists():
        return True, ""
    err = (
        f"tentativa 1\n"
        f"cmd: {' '.join(cmd)}\n"
        f"exit_code: {proc.returncode}\n"
        f"--- STDERR ---\n{(proc.stderr or '').strip()}\n"
        f"--- STDOUT ---\n{(proc.stdout or '').strip()}\n"
        f"--------------"
    )
    out_pdf.unlink(missing_ok=True)
    return False, err


def _sha256_arquivo(path: Path) -> str:
    h = hashlib.sha256()
    with open(_win_long_path(path), "rb") as f:
        while True:
            b = f.read(1024 * 1024)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def _resumo_pdf(path: Path) -> str:
    try:
        size = os.path.getsize(_win_long_path(path))
    except Exception as exc:
        return f"size=erro({exc!r})"

    cab = b""
    try:
        with open(_win_long_path(path), "rb") as f:
            cab = f.read(16)
    except Exception:
        pass

    cab_hex = cab.hex(" ").upper() if cab else "SEM_LEITURA"

    sha = ""
    try:
        sha = _sha256_arquivo(path)
    except Exception as exc:
        sha = f"erro({exc!r})"
    return f"size={size} sha256={sha} header16={cab_hex}"


def _win_long_path(path: Path) -> str:
    s = str(path)
    if s.startswith("\\\\?\\"):
        return s
    if s.startswith("\\\\"):
        return "\\\\?\\UNC\\" + s.lstrip("\\")
    return "\\\\?\\" + s


def _anchor_seguro(anchor: str) -> str:
    raw = (anchor or "").strip()
    if raw.startswith("\\\\"):
        return "UNC_" + raw.strip("\\").replace("\\", "_").replace("/", "_")
    return (raw.replace(":", "").replace("\\", "").replace("/", "") or "SEM_RAIZ")


def _destino_backup_caminho_completo(arquivo: Path, raiz_backup: Path) -> Path:
    p = PureWindowsPath(str(arquivo))
    anchor = p.anchor
    parts = list(p.parts)
    if parts and parts[0] == anchor:
        parts = parts[1:]
    return raiz_backup / _anchor_seguro(anchor) / Path(*parts)


def copiar_backup(arquivo: Path, raiz_backup: Path) -> Path:
    destino = _destino_backup_caminho_completo(arquivo, raiz_backup)
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(_win_long_path(arquivo), _win_long_path(destino))
    return destino


def backup_esta_integro(origem: Path, backup: Path) -> tuple[bool, str]:
    if not backup.exists():
        return False, "arquivo de backup nao existe"

    try:
        size_origem = os.path.getsize(_win_long_path(origem))
        size_backup = os.path.getsize(_win_long_path(backup))
    except Exception as exc:
        return False, f"erro ao ler tamanho: {exc}"

    if size_origem != size_backup:
        return False, f"tamanho diferente (origem={size_origem}, backup={size_backup})"

    try:
        hash_origem = _sha256_arquivo(origem)
        hash_backup = _sha256_arquivo(backup)
    except Exception as exc:
        return False, f"erro ao calcular hash: {exc}"

    if hash_origem != hash_backup:
        return False, "hash SHA-256 diferente"
    return True, ""


def processar_pasta(
    pasta_origem: Path,
    conv_exe: str,
    engine: str,
    min_ratio: float,
    ignorar_validacao_texto: bool,
    pasta_backup_base: Path,
) -> None:
    pdfs = sorted(pasta_origem.rglob("*.pdf"))
    if not pdfs:
        print("[info] Nenhum PDF encontrado para converter.")
        return

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pasta_backup = pasta_backup_base / stamp
    pasta_backup.mkdir(parents=True, exist_ok=True)

    log_path = SCRIPT_DIR / f"log_conversao_pdfa_{stamp}.txt"
    diagnostico_dir = SCRIPT_DIR / f"diagnostico_falhas_{stamp}"
    diagnostico_count = 0
    ok = 0
    skip = 0
    fail = 0

    with log_path.open("w", encoding="utf-8") as log:
        _log(log, f"Origem: {pasta_origem}\n")
        _log(log, f"Backup: {pasta_backup}\n")
        _log(log, f"Engine: {engine}\n")
        _log(log, f"Executavel: {conv_exe}\n")
        if engine == "ghostscript":
            _log(log, f"Ghostscript versao:\n{_gs_versao(conv_exe)}\n")
        _log(log, f"Min ratio texto: {min_ratio}\n")
        _log(log, f"Ignorar validacao de texto: {ignorar_validacao_texto}\n\n")

        _log(log, "FASE 1 - BACKUP E VALIDACAO\n")
        backups_confirmados: dict[Path, Path] = {}
        erro_backup = False

        for pdf in pdfs:
            print(f"[info] Backup: {pdf}")
            _log(log, f"[BACKUP] {pdf}\n")
            try:
                backup_pdf = copiar_backup(pdf, pasta_backup)
                ok_backup, msg_backup = backup_esta_integro(pdf, backup_pdf)
                if not ok_backup:
                    erro_backup = True
                    print(f"[fail] backup invalido: {pdf}")
                    _log(log, f"[FAIL] backup invalido. motivo={msg_backup}\n")
                    continue

                backups_confirmados[pdf] = backup_pdf
                _log(log, f"[OK] backup confirmado: {backup_pdf}\n")
            except Exception as exc:
                erro_backup = True
                print(f"[fail] erro no backup: {pdf}")
                _log(log, f"[FAIL] erro ao criar/validar backup: {exc!r}\n")

        _log(log, "\n")
        if erro_backup or len(backups_confirmados) != len(pdfs):
            print("[erro] Backup nao foi confirmado para todos os arquivos. Conversao cancelada.")
            print(f"[info] Backup parcial em: {pasta_backup}")
            print(f"[info] Log em: {log_path}")
            _log(log, "[ERRO] Conversao cancelada por falha na fase de backup.\n")
            return

        _log(log, "FASE 2 - CONVERSAO DOS ORIGINAIS\n")
        for pdf in pdfs:
            print(f"[info] Convertendo original: {pdf}")
            _log(log, f"[CONVERTER] {pdf}\n")
            _log(log, f"[PDF_ENTRADA] {_resumo_pdf(pdf)}\n")

            backup_pdf = backups_confirmados[pdf]
            with tempfile.TemporaryDirectory(prefix="pdfa_conv_") as td:
                tmp_dir = Path(td)
                src_local = tmp_dir / "origem.pdf"
                out_local = tmp_dir / "convertido_pdfa.pdf"

                try:
                    shutil.copy2(_win_long_path(pdf), src_local)
                except Exception as exc:
                    fail += 1
                    print(f"[fail] erro ao copiar para temporario: {pdf}")
                    _log(log, f"[FAIL] copia para temporario. erro={exc!r}\n\n")
                    continue

                if engine == "ghostscript":
                    convertido, erro = converter_pdfa_ghostscript(conv_exe, src_local, out_local)
                else:
                    convertido, erro = converter_pdfa_ocrmypdf(conv_exe, src_local, out_local)
                if not convertido:
                    src_norm = tmp_dir / "origem_normalizada_pypdf.pdf"
                    ok_norm, err_norm = _normalizar_pdf_com_pypdf(src_local, src_norm)
                    if ok_norm and src_norm.exists():
                        if engine == "ghostscript":
                            convertido2, erro2 = converter_pdfa_ghostscript(conv_exe, src_norm, out_local)
                        else:
                            convertido2, erro2 = converter_pdfa_ocrmypdf(conv_exe, src_norm, out_local)
                        if convertido2:
                            convertido = True
                            erro = ""
                            _log(log, "[INFO] conversao OK apos normalizacao com pypdf.\n")
                        else:
                            erro = (
                                f"{erro}\n\n[FALLBACK_PYPDF] normalizacao executada, mas conversao ainda falhou.\n{erro2}"
                            )
                    else:
                        erro = f"{erro}\n\n[FALLBACK_PYPDF] nao foi possivel normalizar: {err_norm}"

                if not convertido:
                    fail += 1
                    print(f"[fail] erro de conversao: {pdf}")
                    _log(log, f"[FAIL] conversao.\n{erro}\n")
                    if diagnostico_count < MAX_ARQUIVOS_DIAGNOSTICO:
                        diagnostico_count += 1
                        diagnostico_dir.mkdir(parents=True, exist_ok=True)
                        diag_file = diagnostico_dir / f"falha_{diagnostico_count:03d}.pdf"
                        diag_txt = diagnostico_dir / f"falha_{diagnostico_count:03d}.txt"
                        try:
                            shutil.copy2(src_local, diag_file)
                            diag_txt.write_text(
                                f"ORIGINAL: {pdf}\nBACKUP: {backup_pdf}\n{_resumo_pdf(src_local)}\n\n{erro}\n",
                                encoding="utf-8",
                            )
                            _log(log, f"[DIAG] arquivo salvo em: {diag_file}\n")
                            _log(log, f"[DIAG] detalhes salvos em: {diag_txt}\n")
                        except Exception as exc:
                            _log(log, f"[DIAG_FAIL] nao consegui salvar diagnostico: {exc!r}\n")
                    _log(log, "\n")
                    continue

                if not ignorar_validacao_texto:
                    try:
                        txt_entrada = extrair_texto_len(pdf)
                        txt_saida = extrair_texto_len(out_local)
                    except Exception as exc:
                        fail += 1
                        print(f"[fail] validacao indisponivel: {pdf}")
                        _log(log, f"[FAIL] validacao indisponivel. erro={exc}\n\n")
                        continue

                    if txt_entrada > 0:
                        ratio = (txt_saida / txt_entrada) if txt_saida >= 0 else 0.0
                        if ratio < min_ratio:
                            skip += 1
                            print(f"[skip] possivel perda de texto (ratio={ratio:.3f}): {pdf}")
                            _log(log,
                                f"[SKIP] possivel perda de texto. entrada={txt_entrada} saida={txt_saida} ratio={ratio:.3f}\n\n"
                            )
                            continue
                        _log(log,
                            f"[OK] validacao texto. entrada={txt_entrada} saida={txt_saida} ratio={ratio:.3f}\n"
                        )
                    else:
                        _log(log, "[INFO] entrada sem texto extraivel; mantida conversao sem comparacao.\n")

                try:
                    shutil.copy2(out_local, _win_long_path(pdf))
                    ok += 1
                    print(f"[ok] convertido e substituido: {pdf}")
                    _log(log, f"[OK] substituido. backup={backup_pdf}\n\n")
                except Exception as exc:
                    fail += 1
                    print(f"[fail] nao consegui substituir arquivo: {pdf}")
                    _log(log, f"[FAIL] substituir arquivo. erro={exc!r}\n\n")
                    continue

        resumo = f"Resumo: OK={ok} SKIP={skip} FAIL={fail}\n"
        print(resumo.strip())
        print(f"[info] Backup em: {pasta_backup}")
        if diagnostico_dir.exists():
            print(f"[info] Diagnostico de falhas em: {diagnostico_dir}")
        print(f"[info] Log em: {log_path}")
        _log(log, resumo)
        if diagnostico_dir.exists():
            _log(log, f"Diagnostico de falhas: {diagnostico_dir}\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Converte PDFs de uma pasta para PDF/A, substitui no local e cria backup."
    )
    parser.add_argument(
        "--pasta",
        required=True,
        help="Pasta raiz com PDFs para converter (recursivo).",
    )
    parser.add_argument(
        "--gs",
        default="",
        help="Caminho do gswin64c.exe. Se vazio, tenta detectar automaticamente.",
    )
    parser.add_argument(
        "--engine",
        choices=["ghostscript", "ocrmypdf"],
        default="ghostscript",
        help="Motor de conversao PDF/A.",
    )
    parser.add_argument(
        "--ocrmypdf",
        default="",
        help="Caminho do ocrmypdf(.exe). Usado quando --engine ocrmypdf.",
    )
    parser.add_argument(
        "--min-ratio",
        type=float,
        default=0.98,
        help="Razao minima de texto extraivel (saida/entrada). Padrao: 0.98",
    )
    parser.add_argument(
        "--ignorar-validacao-texto",
        action="store_true",
        help="Desativa comparacao de texto extraivel (nao recomendado).",
    )
    parser.add_argument(
        "--pasta-backup",
        default=str(BACKUP_BASE_PADRAO),
        help="Pasta base para salvar backups. Padrao: Desktop\\\\Backups RH",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    pasta = Path(args.pasta).expanduser().resolve()
    if not pasta.exists() or not pasta.is_dir():
        print(f"[erro] pasta invalida: {pasta}")
        sys.exit(1)

    pasta_backup_base = Path(args.pasta_backup).expanduser().resolve()
    pasta_backup_base.mkdir(parents=True, exist_ok=True)

    try:
        if args.engine == "ghostscript":
            conv_exe = localizar_ghostscript(args.gs.strip() or None)
        else:
            conv_exe = localizar_ocrmypdf(args.ocrmypdf.strip() or None)
    except Exception as exc:
        print(f"[erro] {exc}")
        sys.exit(1)

    try:
        processar_pasta(
            pasta_origem=pasta,
            conv_exe=conv_exe,
            engine=args.engine,
            min_ratio=args.min_ratio,
            ignorar_validacao_texto=args.ignorar_validacao_texto,
            pasta_backup_base=pasta_backup_base,
        )
    except KeyboardInterrupt:
        print("[info] Execucao interrompida pelo usuario.")
        sys.exit(130)


if __name__ == "__main__":
    main()
