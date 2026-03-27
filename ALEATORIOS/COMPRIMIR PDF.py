import os
import subprocess
from pathlib import Path

BASE_DIR = r"C:\Users\Usuario\Desktop\pontos"
GS = r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe"  # ajuste se o caminho for outro
REDUCTION_TARGET = 0.01  # só substitui se reduzir 30% ou mais

def log(msg): print(msg, flush=True)

def compress_with_gs(input_pdf: Path, output_pdf: Path) -> bool:
    args = [
        GS,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        "-dPDFSETTINGS=/printer",      # pode testar também /ebook ou /screen
        "-dDetectDuplicateImages=true",
        "-dCompressFonts=true",
        "-dSubsetFonts=true",
        "-dNOPAUSE", "-dQUIET", "-dBATCH",
        f"-sOutputFile={output_pdf}",
        str(input_pdf)
    ]
    try:
        result = subprocess.run(args, capture_output=True, text=True)
        if result.returncode != 0:
            log(f"❌ Erro Ghostscript: {result.stderr.strip()}")
            return False
        return True
    except FileNotFoundError:
        log("❌ Ghostscript não encontrado.")
        return False

def process_pdf(pdf_path: Path):
    log(f"🧩 Processando: {pdf_path}")
    orig_size = pdf_path.stat().st_size
    tmp = pdf_path.with_name(pdf_path.stem + "_tmp.pdf")

    ok = compress_with_gs(pdf_path, tmp)
    if not ok or not tmp.exists():
        return

    new_size = tmp.stat().st_size
    reduc = 1 - new_size / orig_size
    reduc_pct = round(reduc * 100, 2)

    if reduc >= REDUCTION_TARGET:
        tmp.replace(pdf_path)
        log(f"✅ Reduzido {reduc_pct}% — sobrescrito.")
    else:
        tmp.unlink(missing_ok=True)
        log(f"⚠️ Redução {reduc_pct}% (<{int(REDUCTION_TARGET*100)}%). Mantido.")

def main():
    for root, _, files in os.walk(BASE_DIR):
        for f in files:
            if f.lower().endswith(".pdf"):
                process_pdf(Path(root) / f)
    log("🎯 Concluído.")

if __name__ == "__main__":
    main()
