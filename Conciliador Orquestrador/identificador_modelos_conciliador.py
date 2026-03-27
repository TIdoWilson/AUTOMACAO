from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from unicodedata import normalize

from pypdf import PdfReader


MODELO_DESCONHECIDO = "DESCONHECIDO"
MODELO_LIVRO_IOB = "LIVRO_IOB"
MODELO_TIPO50_MAGNETICO = "TIPO50_MAGNETICO"
MODELO_RELACAO_CONFERENCIA_DIA = "RELACAO_CONFERENCIA_DIA"

ORQUESTRADOR_DESCONHECIDO = "DESCONHECIDO"
ORQUESTRADOR_CARTAO_TIPO50 = "CONCILIADOR_CARTAO_TIPO50"
ORQUESTRADOR_PERIN_RELACAO = "CONCILIADOR_RELACAO_DIA"


@dataclass(frozen=True)
class IdentificacaoPDF:
    arquivo: Path
    modelo: str
    movimento: str | None
    score: int
    observacao: str


@dataclass(frozen=True)
class IdentificacaoPar:
    arquivo_a: IdentificacaoPDF
    arquivo_b: IdentificacaoPDF
    orquestrador_sugerido: str
    justificativa: str


def _normalizar_txt(s: str) -> str:
    return normalize("NFKD", s).encode("ASCII", "ignore").decode().upper()


def _extrair_texto_normalizado(pdf_path: Path, max_pages: int = 3) -> str:
    leitor = PdfReader(str(pdf_path))
    partes: list[str] = []
    for pagina in leitor.pages[:max_pages]:
        partes.append(pagina.extract_text() or "")
    return _normalizar_txt("\n".join(partes))


def _contar_termos(texto: str, termos: Iterable[str]) -> int:
    total = 0
    for termo in termos:
        if termo in texto:
            total += 1
    return total


def identificar_pdf(pdf_path: Path) -> IdentificacaoPDF:
    texto = _extrair_texto_normalizado(pdf_path)
    texto_sem_espaco = "".join(texto.split())

    score_livro = _contar_termos(
        texto_sem_espaco,
        (
            "REGISTRODEENTRADAS-MODELOP1",
            "REGISTRODESAIDAS-MODELOP2",
            "COD.DEVALORESFISCAIS",
            "DOCUMENTOSFISCAISCODIFICACAOVALORESFISCAIS",
        ),
    )
    score_tipo50 = _contar_termos(
        texto_sem_espaco,
        (
            "RELATORIODOARQUIVOMAGNETICO-TIPO50",
            "C.F.O.",
            "REGISTROS.:",
            "SITUACAONF:NAOCANCELADA",
        ),
    )
    score_relacao = _contar_termos(
        texto_sem_espaco,
        (
            "RELACAODENOTASPARACONFERENCIAPORDIA",
            "REGISTRODEENTRADAS",
            "REGISTRODESAIDAS",
            "NOTACLIENTEUFVALORIPI",
        ),
    )

    movimento: str | None = None
    if "REGISTRODEENTRADAS" in texto_sem_espaco or "NF:ENTRADAS" in texto_sem_espaco:
        movimento = "ENTRADAS"
    elif "REGISTRODESAIDAS" in texto_sem_espaco or "NF:SAIDAS" in texto_sem_espaco:
        movimento = "SAIDAS"

    placar = {
        MODELO_LIVRO_IOB: score_livro,
        MODELO_TIPO50_MAGNETICO: score_tipo50,
        MODELO_RELACAO_CONFERENCIA_DIA: score_relacao,
    }
    modelo = max(placar, key=placar.get)
    score = placar[modelo]

    if score <= 0:
        modelo = MODELO_DESCONHECIDO
        score = 0
        obs = "Nenhuma assinatura forte reconhecida."
    else:
        obs = f"Assinaturas: livro={score_livro}, tipo50={score_tipo50}, relacao={score_relacao}"

    return IdentificacaoPDF(
        arquivo=pdf_path,
        modelo=modelo,
        movimento=movimento,
        score=score,
        observacao=obs,
    )


def identificar_par(pdf_a: Path, pdf_b: Path) -> IdentificacaoPar:
    a = identificar_pdf(pdf_a)
    b = identificar_pdf(pdf_b)

    modelos = {a.modelo, b.modelo}
    movimentos = {m for m in (a.movimento, b.movimento) if m}

    if MODELO_DESCONHECIDO in modelos:
        return IdentificacaoPar(a, b, ORQUESTRADOR_DESCONHECIDO, "Um ou mais PDFs sem identificação.")

    if len(movimentos) > 1:
        return IdentificacaoPar(a, b, ORQUESTRADOR_DESCONHECIDO, "PDFs de movimentos diferentes.")

    if MODELO_LIVRO_IOB in modelos and MODELO_TIPO50_MAGNETICO in modelos:
        return IdentificacaoPar(
            a,
            b,
            ORQUESTRADOR_CARTAO_TIPO50,
            "Par identificado como Livro IOB + Relatório Tipo 50 magnético.",
        )

    if MODELO_LIVRO_IOB in modelos and MODELO_RELACAO_CONFERENCIA_DIA in modelos:
        return IdentificacaoPar(
            a,
            b,
            ORQUESTRADOR_PERIN_RELACAO,
            "Par identificado como Livro IOB + Relação de Notas por Dia.",
        )

    return IdentificacaoPar(a, b, ORQUESTRADOR_DESCONHECIDO, "Combinação ainda não mapeada.")


def _imprimir_id_pdf(id_pdf: IdentificacaoPDF) -> None:
    print(f"ARQUIVO: {id_pdf.arquivo}")
    print(f"  MODELO: {id_pdf.modelo}")
    print(f"  MOVIMENTO: {id_pdf.movimento}")
    print(f"  SCORE: {id_pdf.score}")
    print(f"  OBS: {id_pdf.observacao}")


def main() -> None:
    import sys

    if len(sys.argv) < 3:
        print("Uso: python identificador_modelos_conciliador.py <pdf_a> <pdf_b>")
        return

    pdf_a = Path(sys.argv[1])
    pdf_b = Path(sys.argv[2])
    if not pdf_a.exists() or not pdf_b.exists():
        print("Erro: caminho de PDF invalido.")
        return

    resultado = identificar_par(pdf_a, pdf_b)
    _imprimir_id_pdf(resultado.arquivo_a)
    _imprimir_id_pdf(resultado.arquivo_b)
    print("ORQUESTRADOR SUGERIDO:", resultado.orquestrador_sugerido)
    print("JUSTIFICATIVA:", resultado.justificativa)


if __name__ == "__main__":
    main()

