from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import sys


@dataclass(frozen=True)
class Tools:
    seven_zip: Path
    gswin64c: Path


class Logger:
    def __init__(self, log_path: Path) -> None:
        self.log_path = log_path
        self.log_path.parent.mkdir(parents=True, exist_ok=True)

    def msg(self, text: str = "") -> None:
        print(text, flush=True)
        with self.log_path.open("a", encoding="utf-8") as f:
            f.write(text + "\n")


def _safe_name(s: str) -> str:
    bad = '<>:"/\\|?*'
    for ch in bad:
        s = s.replace(ch, "_")
    s = " ".join(s.split()).strip().strip(". ")
    return s or "PASTA"


def _folder_label(folder: Path) -> str:
    if folder.name:
        return _safe_name(folder.name)

    anchor = str(folder.anchor).replace(":", "").replace("\\", "").replace("/", "")
    return _safe_name(anchor or "PASTA")


def _is_relative_to(path: Path, possible_parent: Path) -> bool:
    try:
        path.resolve().relative_to(possible_parent.resolve())
        return True
    except Exception:
        return False


def _pick_folder_ui() -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    folder = filedialog.askdirectory(title="Selecione a pasta com PDFs para converter para PDF/A")
    if not folder:
        return None

    try:
        messagebox.showinfo("Selecionado", f"Pasta selecionada:\n{folder}")
    except Exception:
        pass

    return Path(folder)


def _find_7z() -> Path | None:
    found = shutil.which("7z.exe") or shutil.which("7z")
    if found:
        return Path(found)

    candidates = [
        Path(r"C:\Program Files\7-Zip\7z.exe"),
        Path(r"C:\Program Files (x86)\7-Zip\7z.exe"),
    ]
    for c in candidates:
        if c.exists():
            return c

    return None


def _find_ghostscript() -> Path | None:
    found = shutil.which("gswin64c.exe") or shutil.which("gswin32c.exe") or shutil.which("gs")
    if found:
        return Path(found)

    candidates = [
        Path(r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.05.0\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.03.1\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.02.0\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.01.2\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.01.1\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.01.0\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs10.00.0\bin\gswin64c.exe"),
        Path(r"C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.05.1\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.05.0\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.04.0\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.03.1\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.03.0\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.02.1\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.02.0\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.01.2\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.01.1\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.01.0\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs10.00.0\bin\gswin32c.exe"),
        Path(r"C:\Program Files (x86)\gs\gs9.56.1\bin\gswin32c.exe"),
    ]

    for c in candidates:
        if c.exists():
            return c

    return None


def _run(cmd: list[str], cwd: Path | None = None, timeout: int | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        cmd,
        cwd=str(cwd) if cwd else None,
        capture_output=True,
        text=True,
        timeout=timeout,
    )


def _run_checked(cmd: list[str], cwd: Path | None = None, timeout: int | None = None) -> subprocess.CompletedProcess[str]:
    p = _run(cmd, cwd=cwd, timeout=timeout)
    if p.returncode != 0:
        stderr = (p.stderr or "").strip()
        stdout = (p.stdout or "").strip()
        tail = "\n".join([x for x in [stdout, stderr] if x])
        raise RuntimeError(f"Falha ao executar:\n{' '.join(cmd)}\n\nSaída:\n{tail}")
    return p


def _check_tools() -> Tools:
    seven_zip = _find_7z()
    if not seven_zip:
        raise RuntimeError(
            "7-Zip não encontrado.\n"
            "Instale o 7-Zip e garanta o executável '7z.exe'.\n"
            "Caminhos comuns:\n"
            "- C:\\Program Files\\7-Zip\\7z.exe\n"
            "- C:\\Program Files (x86)\\7-Zip\\7z.exe"
        )

    gswin64c = _find_ghostscript()
    if not gswin64c:
        raise RuntimeError(
            "Ghostscript não encontrado.\n"
            "Instale o Ghostscript e garanta o executável 'gswin64c.exe' ou 'gswin32c.exe'.\n"
            "Caminho comum:\n"
            "- C:\\Program Files\\gs\\gsXX.XX.X\\bin\\gswin64c.exe"
        )

    _run_checked([str(seven_zip), "i"])
    _run_checked([str(gswin64c), "-version"])

    return Tools(seven_zip=seven_zip, gswin64c=gswin64c)


def _iter_pdfs(root: Path) -> list[Path]:
    pdfs: list[Path] = []
    for p in root.rglob("*"):
        if p.is_file() and p.suffix.lower() == ".pdf":
            pdfs.append(p)

    pdfs.sort(key=lambda x: str(x).lower())
    return pdfs


def _choose_backup_root(target_dir: Path, script_dir: Path) -> Path:
    preferred = script_dir.parent / "_BACKUPS_PDF_A"
    if _is_relative_to(preferred, target_dir) or preferred.resolve() == target_dir.resolve():
        preferred = Path.home() / "_BACKUPS_PDF_A"
    return preferred


def _choose_log_root(target_dir: Path, script_dir: Path) -> Path:
    preferred = script_dir.parent / "_LOGS_PDF_A"
    if _is_relative_to(preferred, target_dir) or preferred.resolve() == target_dir.resolve():
        preferred = Path.home() / "_LOGS_PDF_A"
    return preferred


def _choose_output_dir(target_dir: Path, script_dir: Path) -> Path:
    base = _folder_label(target_dir)
    candidate = target_dir.parent / f"{base}__PDF_A"

    if _is_relative_to(candidate, target_dir) or candidate.resolve() == target_dir.resolve():
        candidate = script_dir.parent / f"{base}__PDF_A"

    if _is_relative_to(candidate, target_dir) or candidate.resolve() == target_dir.resolve():
        candidate = Path.home() / f"{base}__PDF_A"

    return candidate


def _make_backup_7z(tools: Tools, target_dir: Path, backup_root: Path, logger: Logger) -> Path:
    backup_root.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = _folder_label(target_dir)
    out_7z = backup_root / f"BACKUP_{base}_{stamp}.7z"

    logger.msg(f"[BACKUP] Criando: {out_7z}")
    _run_checked([str(tools.seven_zip), "a", "-t7z", str(out_7z), str(target_dir)])

    if not out_7z.exists():
        raise RuntimeError(f"Backup não foi criado: {out_7z}")

    size = out_7z.stat().st_size
    if size <= 0:
        raise RuntimeError(f"Backup criado com tamanho inválido: {out_7z}")

    logger.msg(f"[BACKUP] Validando integridade: {out_7z}")
    _run_checked([str(tools.seven_zip), "t", str(out_7z)])

    logger.msg(f"[BACKUP] Integridade OK | Tamanho: {size:,} bytes")
    return out_7z


def _convert_one_pdf(tools: Tools, src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)

    tmp = dst.with_name(dst.stem + ".__tmp_pdfa__.pdf")
    if tmp.exists():
        tmp.unlink()

    if dst.exists():
        dst.unlink()

    # Conversão para PDF/A sem OCR e sem adicionar camada de texto.
    # Preserva o texto já existente; apenas regrava em PDF/A.
    cmd = [
        str(tools.gswin64c),
        "-dPDFA=2",
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        "-dNOSAFER",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.7",
        "-dAutoRotatePages=/None",
        "-dEmbedAllFonts=true",
        "-dSubsetFonts=true",
        "-dCompressFonts=true",
        f"-sOutputFile={str(tmp)}",
        str(src),
    ]

    try:
        _run_checked(cmd, timeout=900)

        if not tmp.exists():
            raise RuntimeError("Ghostscript terminou sem gerar arquivo de saída.")

        if tmp.stat().st_size <= 0:
            raise RuntimeError("Arquivo de saída PDF/A foi gerado com tamanho zero.")

        tmp.replace(dst)
    except Exception:
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass
        raise


def convert_folder(tools: Tools, target_dir: Path, out_dir: Path, pdfs: list[Path], logger: Logger) -> tuple[int, int]:
    if not pdfs:
        logger.msg("[INFO] Nenhum PDF encontrado na pasta selecionada.")
        return 0, 0

    ok = 0
    fail = 0

    logger.msg(f"[INFO] PDFs encontrados: {len(pdfs)}")
    logger.msg(f"[INFO] Saída PDF/A: {out_dir}")
    logger.msg("")

    for i, src in enumerate(pdfs, start=1):
        rel = src.relative_to(target_dir)
        dst = out_dir / rel

        logger.msg(f"[{i}/{len(pdfs)}] Convertendo: {rel}")
        try:
            _convert_one_pdf(tools, src, dst)
            ok += 1
            logger.msg(f"  [OK] {rel}")
        except Exception as e:
            fail += 1
            logger.msg(f"  [ERRO] {rel}: {e}")
        logger.msg("")

    return ok, fail


def _get_target_dir_from_args_or_ui() -> Path | None:
    if len(sys.argv) >= 2:
        return Path(sys.argv[1])
    return _pick_folder_ui()


def main() -> None:
    script_dir = Path(__file__).resolve().parent
    target_dir = _get_target_dir_from_args_or_ui()

    if not target_dir:
        print("Nenhuma pasta selecionada. Cancelado.", flush=True)
        return

    if not target_dir.exists() or not target_dir.is_dir():
        print("Pasta inválida. Cancelado.", flush=True)
        return

    backup_root = _choose_backup_root(target_dir, script_dir)
    log_root = _choose_log_root(target_dir, script_dir)
    out_dir = _choose_output_dir(target_dir, script_dir)

    log_root.mkdir(parents=True, exist_ok=True)
    log_path = log_root / f"PDF_A_{_folder_label(target_dir)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logger = Logger(log_path)

    logger.msg("")
    logger.msg(f"[INFO] Pasta alvo: {target_dir}")
    logger.msg(f"[INFO] Pasta de logs: {log_root}")
    logger.msg(f"[INFO] Pasta de backup: {backup_root}")
    logger.msg(f"[INFO] Pasta de saída: {out_dir}")
    logger.msg("")

    pdfs = _iter_pdfs(target_dir)
    if not pdfs:
        logger.msg("[INFO] Nenhum PDF encontrado na pasta selecionada.")
        logger.msg(f"[INFO] Log: {log_path}")
        return

    try:
        tools = _check_tools()
    except Exception as e:
        logger.msg(str(e))
        logger.msg("")
        logger.msg("Dica de instalação:")
        logger.msg("  choco install 7zip -y")
        logger.msg("  choco install ghostscript -y")
        logger.msg("")
        logger.msg(f"[INFO] Log: {log_path}")
        return

    logger.msg(f"[INFO] 7-Zip: {tools.seven_zip}")
    logger.msg(f"[INFO] Ghostscript: {tools.gswin64c}")
    logger.msg(f"[INFO] PDFs localizados: {len(pdfs)}")
    logger.msg("")

    try:
        backup_path = _make_backup_7z(tools, target_dir, backup_root, logger)
    except Exception as e:
        logger.msg("[FATAL] Backup falhou. Conversão abortada.")
        logger.msg(f"[FATAL] Motivo: {e}")
        logger.msg("")
        logger.msg(f"[INFO] Log: {log_path}")
        return

    logger.msg(f"[BACKUP] OK: {backup_path}")
    logger.msg("")

    ok, fail = convert_folder(tools, target_dir, out_dir, pdfs, logger)

    logger.msg("======================================")
    logger.msg(f"Concluído. OK={ok} | FALHAS={fail}")
    logger.msg(f"Backup validado: {backup_path}")
    logger.msg(f"Saída PDF/A: {out_dir}")
    logger.msg(f"Log: {log_path}")
    logger.msg("======================================")


if __name__ == "__main__":
    main()