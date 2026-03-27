import fitz  # PyMuPDF
from pathlib import Path
import traceback
import shutil

# === CONFIGURAÇÕES ===
BASE_DIR = Path(r"C:\Users\Usuario\Pictures\Saved Pictures\pontos")
JPEG_QUALITY = 50        # 40–60 = boa compressão
DPI_SCALE = 1.0          # 1.0 = ~72dpi, 1.5 = melhor nitidez
# ======================

def log(msg): print(msg, flush=True)

def convert_pdf_to_bw(pdf_path: Path) -> Path | None:
    """Converte o PDF para tons de cinza e retorna o caminho do novo arquivo."""
    try:
        log(f"🧩 Convertendo: {pdf_path}")
        doc = fitz.open(pdf_path)
        out = fitz.open()

        for page in doc:
            # Renderiza página em tons de cinza
            pix = page.get_pixmap(matrix=fitz.Matrix(DPI_SCALE, DPI_SCALE), colorspace=fitz.csGRAY)
            img_bytes = pix.tobytes("jpeg", jpg_quality=JPEG_QUALITY)

            # Nova página no PDF de saída
            width_pt, height_pt = pix.width, pix.height
            new_page = out.new_page(width=width_pt, height=height_pt)
            rect = fitz.Rect(0, 0, width_pt, height_pt)
            new_page.insert_image(rect, stream=img_bytes)

        # Arquivo temporário no mesmo diretório
        temp_pdf = pdf_path.with_name(pdf_path.stem + "_tmp.pdf")
        out.save(temp_pdf, deflate=True)
        out.close()
        doc.close()
        return temp_pdf

    except Exception as e:
        log(f"❌ Erro ao converter {pdf_path}: {e}")
        traceback.print_exc()
        return None

def process_pdf(pdf_path: Path):
    try:
        orig_size = pdf_path.stat().st_size
        temp_pdf = convert_pdf_to_bw(pdf_path)
        if not temp_pdf or not temp_pdf.exists():
            log("  ❌ Falha ao gerar PDF temporário.\n")
            return

        new_size = temp_pdf.stat().st_size
        reduc = (1 - new_size / orig_size) * 100 if orig_size else 0

        if new_size < orig_size:
            # Sobrescreve com backup
            backup = pdf_path.with_suffix(".orig_backup.pdf")
            try:
                shutil.move(str(pdf_path), str(backup))
                shutil.move(str(temp_pdf), str(pdf_path))
                backup.unlink(missing_ok=True)
                log(f"  ✅ Redução {reduc:.1f}% — sobrescrito.\n")
            except Exception as e:
                log(f"  ❌ Falha ao sobrescrever {pdf_path}: {e}\n")
        else:
            log(f"  ⚠️ Sem ganho ({reduc:.1f}%), mantido original.\n")
            temp_pdf.unlink(missing_ok=True)

    except Exception as e:
        log(f"❌ Erro inesperado com {pdf_path}: {e}")
        traceback.print_exc()

def main():
    if not BASE_DIR.exists():
        log(f"❌ Pasta não encontrada: {BASE_DIR}")
        return

    pdfs = list(BASE_DIR.rglob("*.pdf"))
    if not pdfs:
        log("⚠️ Nenhum PDF encontrado.")
        return

    log(f"🔎 Encontrados {len(pdfs)} PDFs em {BASE_DIR}")
    for pdf_path in pdfs:
        process_pdf(pdf_path)

    log("🎯 Conversão concluída.")

if __name__ == "__main__":
    main()
