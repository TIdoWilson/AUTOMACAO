from __future__ import annotations

import argparse
import csv
import re
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import fitz


MONEY_RE = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
PDF_START_RE = re.compile(r"^\d{2}/\d{2}/\d{4}\s+LANC CTA RECEB:\s*(.*)$")
PDF_CP_LOTE_RE = re.compile(r"^\d{1,3}-\d+\s+\d+/\d+$")
DOC_RE = re.compile(r"LANC CTA RECEB:\s*(\d+)-")
TXT_L_RE = re.compile(r"^L(\d{8})\d{15}\s+(\d{18})\s+\d{5}\|")
FS_DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
UF_RE = re.compile(r"^[A-Z]{2}$")


@dataclass(frozen=True)
class Lancamento:
    origem: str
    data: str
    doc: str
    valor: Decimal
    historico: str


def parse_moeda_br(valor: str) -> Decimal:
    valor = valor.strip().replace(".", "").replace(",", ".")
    return Decimal(valor)


def decimal_txt_raw(raw: str) -> Decimal:
    # Valor no TXT vem com 3 casas extras (ex.: 211877000 -> 2.118,77)
    return (Decimal(int(raw)) / Decimal("100000")).quantize(Decimal("0.01"))


def extrair_txt(path: Path) -> list[Lancamento]:
    registros: list[Lancamento] = []
    ultimo_l: tuple[str, Decimal] | None = None

    with path.open("r", encoding="latin-1", errors="ignore") as f:
        for linha in f:
            linha = linha.rstrip("\n")
            m_l = TXT_L_RE.match(linha[:110])
            if m_l:
                data_raw, valor_raw = m_l.groups()
                data = f"{data_raw[0:2]}/{data_raw[2:4]}/{data_raw[4:8]}"
                valor = decimal_txt_raw(valor_raw)
                ultimo_l = (data, valor)
                continue

            if not linha.startswith("H") or ultimo_l is None:
                continue

            historico = linha[1:].strip()
            if "LANC CTA RECEB:" not in historico:
                continue

            m_doc = DOC_RE.search(historico)
            if not m_doc:
                continue

            data, valor = ultimo_l
            registros.append(
                Lancamento(
                    origem="TXT",
                    data=data,
                    doc=m_doc.group(1),
                    valor=valor,
                    historico=historico,
                )
            )
    return registros


def extrair_pdf(path: Path) -> list[Lancamento]:
    registros: list[Lancamento] = []
    doc = fitz.open(path)
    try:
        linhas: list[str] = []
        for pagina in doc:
            linhas.extend(l.strip() for l in pagina.get_text("text").splitlines())

        i = 0
        while i < len(linhas):
            linha = linhas[i]
            m_ini = PDF_START_RE.match(linha)
            if not m_ini:
                i += 1
                continue

            historico = f"LANC CTA RECEB: {m_ini.group(1)}".strip()
            data = linha[:10]
            j = i + 1

            # Junta continuação do histórico até achar linha CP/Lote.
            while j < len(linhas) and not PDF_CP_LOTE_RE.match(linhas[j]):
                if linhas[j]:
                    historico += " " + linhas[j]
                j += 1

            if j >= len(linhas):
                i += 1
                continue

            # Após CP/Lote, pega o primeiro valor monetário (débito/crédito).
            valor: Decimal | None = None
            k = j + 1
            while k < len(linhas):
                atual = linhas[k]
                if MONEY_RE.match(atual):
                    valor = parse_moeda_br(atual)
                    break
                if PDF_START_RE.match(atual):
                    break
                k += 1

            if valor is None:
                i = j + 1
                continue

            m_doc = DOC_RE.search(historico)
            if m_doc:
                registros.append(
                    Lancamento(
                        origem="PDF",
                        data=data,
                        doc=m_doc.group(1),
                        valor=valor,
                        historico=historico,
                    )
                )
            i = k + 1
    finally:
        doc.close()
    return registros


def extrair_pdf_fs002(path: Path) -> list[Lancamento]:
    registros: list[Lancamento] = []
    doc = fitz.open(path)
    try:
        linhas: list[str] = []
        for pagina in doc:
            linhas.extend(l.strip() for l in pagina.get_text("text").splitlines() if l.strip())

        i = 0
        while i < len(linhas):
            linha = linhas[i]
            if not FS_DATE_RE.match(linha):
                i += 1
                continue

            # Layout esperado (FS-002):
            # [numero] [data] [UF] [valor] [numero] [valor] [cliente...]
            if i == 0:
                i += 1
                continue
            numero_antes = linhas[i - 1]
            uf = linhas[i + 1] if i + 1 < len(linhas) else ""
            valor = linhas[i + 2] if i + 2 < len(linhas) else ""
            numero_rep = linhas[i + 3] if i + 3 < len(linhas) else ""
            valor_rep = linhas[i + 4] if i + 4 < len(linhas) else ""

            if (
                numero_antes.isdigit()
                and UF_RE.match(uf)
                and MONEY_RE.match(valor)
                and numero_rep == numero_antes
                and MONEY_RE.match(valor_rep)
                and valor_rep == valor
            ):
                data = linha
                val = parse_moeda_br(valor)
                registros.append(
                    Lancamento(
                        origem="PDF",
                        data=data,
                        doc=numero_antes,
                        valor=val,
                        historico=f"FS002 DOC {numero_antes}",
                    )
                )
                i += 5
                continue
            i += 1
    finally:
        doc.close()
    return registros


def detectar_tipo_pdf(path: Path) -> str:
    doc = fitz.open(path)
    try:
        amostra = "\n".join(doc[i].get_text("text") for i in range(min(2, doc.page_count)))
    finally:
        doc.close()
    if "Registros de Saídas" in amostra or "Registros de Saidas" in amostra:
        return "fs002"
    return "razao"


def formatar_valor(valor: Decimal) -> str:
    s = f"{valor:.2f}"
    inteiro, frac = s.split(".")
    partes = []
    while len(inteiro) > 3:
        partes.insert(0, inteiro[-3:])
        inteiro = inteiro[:-3]
    partes.insert(0, inteiro)
    return f"{'.'.join(partes)},{frac}"


def escrever_csv(path: Path, itens: list[Lancamento]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["origem", "data", "doc", "valor", "historico"])
        for it in itens:
            w.writerow([it.origem, it.data, it.doc, formatar_valor(it.valor), it.historico])


def chave_lancamento(item: Lancamento, modo: str) -> tuple[str, Decimal]:
    valor = item.valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if modo == "data_valor":
        return item.data, valor
    if modo == "doc_valor":
        return item.doc, valor
    raise ValueError(f"Modo de chave invalido: {modo}")


def conciliar(
    txt_items: list[Lancamento], pdf_items: list[Lancamento], modo: str
) -> tuple[list[Lancamento], list[Lancamento], list[Lancamento]]:
    txt_map: dict[tuple[str, Decimal], list[Lancamento]] = defaultdict(list)
    pdf_map: dict[tuple[str, Decimal], list[Lancamento]] = defaultdict(list)
    for x in txt_items:
        txt_map[chave_lancamento(x, modo)].append(x)
    for x in pdf_items:
        pdf_map[chave_lancamento(x, modo)].append(x)

    conciliados: list[Lancamento] = []
    so_txt: list[Lancamento] = []
    so_pdf: list[Lancamento] = []

    todas_chaves = set(txt_map) | set(pdf_map)
    for chave in todas_chaves:
        l_txt = txt_map.get(chave, [])
        l_pdf = pdf_map.get(chave, [])
        n = min(len(l_txt), len(l_pdf))
        if n:
            conciliados.extend(l_txt[:n])
        if len(l_txt) > n:
            so_txt.extend(l_txt[n:])
        if len(l_pdf) > n:
            so_pdf.extend(l_pdf[n:])

    return conciliados, so_txt, so_pdf


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Concilia recebimentos entre arquivo TXT e Razao Analitico em PDF."
    )
    parser.add_argument("--txt", required=True, help="Caminho do arquivo TXT.")
    parser.add_argument("--pdf", required=True, help="Caminho do arquivo PDF.")
    parser.add_argument(
        "--out-dir",
        default=".",
        help="Diretorio de saida dos CSVs. Padrao: diretorio atual.",
    )
    parser.add_argument(
        "--modo-chave",
        choices=["data_valor", "doc_valor"],
        default="data_valor",
        help="Criterio de conciliacao. Padrao: data_valor.",
    )
    parser.add_argument(
        "--pdf-tipo",
        choices=["auto", "razao", "fs002"],
        default="auto",
        help="Tipo do PDF. Padrao: auto.",
    )
    args = parser.parse_args()

    txt_path = Path(args.txt)
    pdf_path = Path(args.pdf)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    txt_items = extrair_txt(txt_path)
    pdf_tipo = args.pdf_tipo if args.pdf_tipo != "auto" else detectar_tipo_pdf(pdf_path)
    if pdf_tipo == "fs002":
        pdf_items = extrair_pdf_fs002(pdf_path)
    else:
        pdf_items = extrair_pdf(pdf_path)
    conciliados, so_txt, so_pdf = conciliar(txt_items, pdf_items, args.modo_chave)

    total_txt = sum((x.valor for x in txt_items), Decimal("0.00"))
    total_pdf = sum((x.valor for x in pdf_items), Decimal("0.00"))
    diferenca = (total_txt - total_pdf).quantize(Decimal("0.01"))
    pendente_txt = sum((x.valor for x in so_txt), Decimal("0.00"))
    pendente_pdf = sum((x.valor for x in so_pdf), Decimal("0.00"))

    escrever_csv(out_dir / "conciliados.csv", conciliados)
    escrever_csv(out_dir / "somente_txt.csv", so_txt)
    escrever_csv(out_dir / "somente_pdf.csv", so_pdf)

    print("=== RESUMO DA CONCILIACAO ===")
    print(f"Tipo de PDF: {pdf_tipo}")
    print(f"Modo de chave: {args.modo_chave}")
    print(f"TXT: {len(txt_items)} lancamentos | Total: {formatar_valor(total_txt)}")
    print(f"PDF: {len(pdf_items)} lancamentos | Total: {formatar_valor(total_pdf)}")
    print(f"Conciliados: {len(conciliados)}")
    print(f"So no TXT: {len(so_txt)} | Total: {formatar_valor(pendente_txt)}")
    print(f"So no PDF: {len(so_pdf)} | Total: {formatar_valor(pendente_pdf)}")
    print(f"Diferenca (TXT - PDF): {formatar_valor(diferenca)}")
    print(f"Arquivos gerados em: {out_dir.resolve()}")


if __name__ == "__main__":
    main()
