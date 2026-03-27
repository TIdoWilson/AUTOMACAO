# extrair_faturas_nf3e.py
# Uso:
#   python extrair_faturas_nf3e.py "caminho/para/pasta_com_pdfs"
#
# Saídas:
#   faturas_extraidas.xlsx
#   faturas_extraidas.csv

import re
import sys
from pathlib import Path

import pdfplumber
import pandas as pd

RE_NOTA = re.compile(
    r"NOTA\s+FISCAL\s+N[oº]\.?\s*(\d+)\s*-\s*S[ÉE]RIE\s*(\d+)\s*/\s*DATA\s+DE\s+EMISS[ÃA]O:\s*(\d{2}/\d{2}/\d{4})",
    re.IGNORECASE,
)

# A chave normalmente tem 44 dígitos (pode vir com espaços/quebras de linha)
RE_CHAVE = re.compile(
    r"Chave\s+de\s+Acesso\s*[\r\n ]*([0-9 \r\n]{40,120})",
    re.IGNORECASE,
)


def extrair_texto_pdf(pdf_path: Path) -> str:
    partes = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            partes.append(page.extract_text() or "")
    return "\n".join(partes)


def parse_campos(texto: str):
    # tenta com o texto bruto e com uma versão "achatada"
    texto_flat = re.sub(r"[ \t]+", " ", texto)

    m = RE_NOTA.search(texto) or RE_NOTA.search(texto_flat)
    numero_nota = serie = data_emissao = None
    if m:
        numero_nota, serie, data_emissao = m.group(1), m.group(2), m.group(3)

    m2 = RE_CHAVE.search(texto) or RE_CHAVE.search(texto_flat)
    chave = None
    if m2:
        apenas_digitos = re.sub(r"\D", "", m2.group(1))
        # se tiver mais que 44 por algum “vazamento”, pega os primeiros 44
        chave = apenas_digitos[:44] if len(apenas_digitos) >= 44 else apenas_digitos

    return data_emissao, numero_nota, serie, chave


def main():
    pasta = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(".")
    pdfs = sorted(pasta.glob("*.pdf"))

    if not pdfs:
        print(f"Nenhum PDF encontrado em: {pasta.resolve()}")
        sys.exit(1)

    linhas = []
    for pdf in pdfs:
        try:
            texto = extrair_texto_pdf(pdf)
            data_emissao, numero_nota, serie, chave = parse_campos(texto)

            erros = []
            if not data_emissao:
                erros.append("data_emissao_nao_encontrada")
            if not numero_nota:
                erros.append("numero_nota_nao_encontrado")
            if not serie:
                erros.append("serie_nao_encontrada")
            if not chave:
                erros.append("chave_nao_encontrada")
            elif len(chave) != 44:
                erros.append(f"chave_tamanho_{len(chave)}")

            linhas.append(
                {
                    "arquivo": pdf.name,
                    "data_emissao": data_emissao,
                    "numero_nota": numero_nota,
                    "serie": serie,
                    "chave_acesso": chave,
                    "erros": ";".join(erros) if erros else "",
                }
            )
        except Exception as e:
            linhas.append(
                {
                    "arquivo": pdf.name,
                    "data_emissao": None,
                    "numero_nota": None,
                    "serie": None,
                    "chave_acesso": None,
                    "erros": f"falha_leitura_pdf:{type(e).__name__}:{e}",
                }
            )

    df = pd.DataFrame(linhas)

    # Salva
    out_xlsx = Path("faturas_extraidas.xlsx")
    out_csv = Path("faturas_extraidas.csv")
    df.to_excel(out_xlsx, index=False)
    df.to_csv(out_csv, index=False, encoding="utf-8-sig")

    print(f"OK: {out_xlsx.resolve()}")
    print(f"OK: {out_csv.resolve()}")
    # mostra rapidamente o que ficou com erro
    com_erro = df[df["erros"].astype(str).str.len() > 0]
    if not com_erro.empty:
        print("\nATENÇÃO: alguns arquivos ficaram com pendências na coluna 'erros':")
        print(com_erro[["arquivo", "erros"]].to_string(index=False))


if __name__ == "__main__":
    main()
