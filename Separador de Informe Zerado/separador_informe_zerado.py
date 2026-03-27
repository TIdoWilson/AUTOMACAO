import argparse
import hashlib
import re
from pathlib import Path

from pypdf import PdfReader, PdfWriter

PADRAO_VALOR = re.compile(r"(?<!\d)(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})(?!\d)")


def extrair_numeros_monetarios(texto: str) -> list[float]:
    valores = PADRAO_VALOR.findall(texto or "")
    numeros: list[float] = []
    for valor in valores:
        try:
            numeros.append(float(valor.replace(".", "").replace(",", ".")))
        except ValueError:
            continue
    return numeros


def pagina_tem_valor_maior_que_zero(texto: str) -> bool:
    return any(n > 0 for n in extrair_numeros_monetarios(texto))


def hash_texto(texto: str) -> str:
    normalizado = " ".join((texto or "").split())
    return hashlib.sha1(normalizado.encode("utf-8", errors="ignore")).hexdigest()


def gerar_pdf_filtrado(origem: Path, saida: Path) -> tuple[list[int], list[int]]:
    leitor = PdfReader(str(origem))
    escritor = PdfWriter()

    paginas_mantidas: list[int] = []
    paginas_removidas: list[int] = []

    for indice, pagina in enumerate(leitor.pages, start=1):
        texto = pagina.extract_text() or ""
        if pagina_tem_valor_maior_que_zero(texto):
            escritor.add_page(pagina)
            paginas_mantidas.append(indice)
        else:
            paginas_removidas.append(indice)

    saida.parent.mkdir(parents=True, exist_ok=True)
    with saida.open("wb") as arquivo_saida:
        escritor.write(arquivo_saida)

    return paginas_mantidas, paginas_removidas


def validar_saida(origem: Path, saida: Path, paginas_esperadas: list[int]) -> tuple[bool, str]:
    leitor_origem = PdfReader(str(origem))
    leitor_saida = PdfReader(str(saida))

    if len(leitor_saida.pages) != len(paginas_esperadas):
        return False, (
            f"Quantidade divergente: saida={len(leitor_saida.pages)} "
            f"esperado={len(paginas_esperadas)}"
        )

    hashes_esperados = []
    for indice in paginas_esperadas:
        texto = leitor_origem.pages[indice - 1].extract_text() or ""
        hashes_esperados.append(hash_texto(texto))

    hashes_saida = []
    for pagina in leitor_saida.pages:
        texto = pagina.extract_text() or ""
        hashes_saida.append(hash_texto(texto))

    if hashes_esperados != hashes_saida:
        return False, "Ordem/conteudo das paginas filtradas nao confere com o esperado."

    return True, "Validacao OK."


def caminho_saida_padrao(origem: Path) -> Path:
    return origem.with_name(f"{origem.stem} - SOMENTE COM VALOR.pdf")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Gera copia do informe mantendo so paginas com valor maior que 0,00."
    )
    parser.add_argument("arquivo_pdf", help="Caminho do PDF de origem")
    parser.add_argument(
        "--saida",
        help="Caminho do PDF de saida (opcional). Se nao informar, usa sufixo padrao.",
    )
    parser.add_argument(
        "--sem-validacao",
        action="store_true",
        help="Nao executar validacao de conteudo apos gerar a copia.",
    )
    args = parser.parse_args()

    origem = Path(args.arquivo_pdf)
    if not origem.exists():
        raise SystemExit(f"Arquivo nao encontrado: {origem}")

    if origem.suffix.lower() != ".pdf":
        raise SystemExit("O arquivo de entrada precisa ser .pdf")

    saida = Path(args.saida) if args.saida else caminho_saida_padrao(origem)

    paginas_mantidas, paginas_removidas = gerar_pdf_filtrado(origem, saida)

    print(f"Origem: {origem}")
    print(f"Saida: {saida}")
    print(f"Total de paginas: {len(paginas_mantidas) + len(paginas_removidas)}")
    print(f"Paginas mantidas: {len(paginas_mantidas)}")
    print(f"Paginas removidas: {len(paginas_removidas)}")

    if paginas_mantidas:
        print("Primeiras paginas mantidas:", ",".join(map(str, paginas_mantidas[:20])))

    if not args.sem_validacao:
        ok, msg = validar_saida(origem, saida, paginas_mantidas)
        print(msg)
        if not ok:
            raise SystemExit(2)


if __name__ == "__main__":
    main()
