from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

try:
    from pypdf import PdfReader
except ImportError as exc:
    print("Erro: modulo 'pypdf' nao encontrado. Instale com: pip install pypdf")
    raise SystemExit(1) from exc


# =========================
# Configuracao
# =========================
CURDEF = "BRL"
ACCTTYPE = "CHECKING"
BANKID_PADRAO = "0000"
ACCTID_PADRAO = "0000"
PASTA_SAIDA_PADRAO = Path(r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\Conversor Extrato PDF para OFX\OFXs Gerados")
OFX_HEADER = """OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE
"""


@dataclass
class Lancamento:
    data: datetime
    valor: float
    memo: str
    documento: str

    @property
    def trntype(self) -> str:
        return "CREDIT" if self.valor >= 0 else "DEBIT"

    @property
    def dtposted(self) -> str:
        return self.data.strftime("%Y%m%d") + "120000"


def parse_valor_br(texto: str) -> float:
    texto = texto.strip()
    negativo = texto.endswith("-")
    if negativo:
        texto = texto[:-1]
    texto = texto.replace(".", "").replace(",", ".")
    valor = float(texto)
    return -valor if negativo else valor


def formatar_valor_br(valor: float) -> str:
    sinal = "-" if valor < 0 else ""
    numero = f"{abs(valor):.2f}".replace(".", ",")
    return f"{sinal}{numero}"


def ler_texto_pdf(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path), strict=False)
    partes: list[str] = []
    for pagina in reader.pages:
        partes.append(pagina.extract_text() or "")
    return "\n".join(partes)


def detectar_banco(texto_pdf: str) -> str:
    texto_maiusculo = texto_pdf.upper()
    if "EVOLUA" in texto_maiusculo and "EXTRATO ESPECIAL" in texto_maiusculo:
        return "evolua"
    if "SICREDI" in texto_maiusculo and "EXTRATO DE CONTA CORRENTE" in texto_maiusculo:
        return "sicredi"
    raise ValueError("Nao foi possivel identificar o banco neste PDF.")


def extrair_conta(texto_pdf: str, banco: str) -> str:
    if banco == "evolua":
        m = re.search(r"CONTA/DV:\s*([0-9\.\-]+)", texto_pdf, flags=re.IGNORECASE)
        if m:
            return re.sub(r"\D", "", m.group(1))
    if banco == "sicredi":
        m = re.search(r"\b(\d{6,}-\d)\b", texto_pdf)
        if m:
            return m.group(1).replace("-", "")
    return ACCTID_PADRAO


def parse_evolua(texto_pdf: str) -> tuple[list[Lancamento], float | None]:
    regex_linha = re.compile(
        r"^\s*(\d{2}/\d{2}/\d{2})\s+(.+?)\s{2,}([0-9 .:/-]+?)\s{2,}([\d\.,-]+)\s+([CD])(?:\s{2,}([\d\.,-]+))?\s*$"
    )
    lancamentos: list[Lancamento] = []
    saldo_final: float | None = None

    for linha in texto_pdf.splitlines():
        m = regex_linha.match(linha)
        if not m:
            continue
        data_txt, historico, documento, valor_txt, dc, saldo_txt = m.groups()
        data = datetime.strptime(data_txt, "%d/%m/%y")
        valor = parse_valor_br(valor_txt)
        if dc == "D":
            valor = -abs(valor)
        else:
            valor = abs(valor)
        if saldo_txt:
            saldo_final = parse_valor_br(saldo_txt)
        lancamentos.append(
            Lancamento(
                data=data,
                valor=valor,
                memo=historico.strip()[:80],
                documento=documento.strip()[:30],
            )
        )
    return lancamentos, saldo_final


def parse_sicredi(texto_pdf: str) -> tuple[list[Lancamento], float | None]:
    regex_base = re.compile(r"^\s*(\d{2}/\d{2}/\d{4})\s+(\S+)\s+(.+)$")
    regex_valor = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}-?")

    lancamentos: list[Lancamento] = []
    saldo_final: float | None = None

    for linha in texto_pdf.splitlines():
        m = regex_base.match(linha)
        if not m:
            continue
        data_txt, documento, resto = m.groups()
        valores = list(regex_valor.finditer(resto))
        if not valores:
            continue

        # Em alguns historicos pode existir valor textual (ex.: "... ARMARINHOS 1,99")
        # antes da coluna real de debito/credito. Prioriza valores com "cara de coluna":
        # precedidos por 2+ espacos.
        valores_coluna: list[re.Match[str]] = []
        for m_valor in valores:
            i = m_valor.start() - 1
            espacos_antes = 0
            while i >= 0 and resto[i] == " ":
                espacos_antes += 1
                i -= 1
            if espacos_antes >= 2:
                valores_coluna.append(m_valor)

        valores_escolhidos = valores_coluna if valores_coluna else valores
        primeiro = valores_escolhidos[0]
        historico = resto[: primeiro.start()].strip()
        valor = parse_valor_br(primeiro.group())

        # No extrato Sicredi de exemplo, debito inicia por volta da coluna 60
        # e credito por volta da coluna 80 no texto extraido.
        eh_credito = primeiro.start() >= 72
        valor = abs(valor) if eh_credito else -abs(valor)

        if len(valores_escolhidos) >= 2:
            saldo_final = parse_valor_br(valores_escolhidos[-1].group())
        elif len(valores) >= 2:
            saldo_final = parse_valor_br(valores[-1].group())

        lancamentos.append(
            Lancamento(
                data=datetime.strptime(data_txt, "%d/%m/%Y"),
                valor=valor,
                memo=historico[:80],
                documento=documento[:30],
            )
        )

    return lancamentos, saldo_final


def montar_ofx(
    lancamentos: Iterable[Lancamento],
    bankid: str,
    acctid: str,
    saldo_final: float | None,
) -> str:
    itens = sorted(lancamentos, key=lambda x: x.data)
    if not itens:
        raise ValueError("Nenhum lancamento foi identificado para gerar OFX.")

    dtstart = itens[0].dtposted
    dtend = itens[-1].dtposted
    if saldo_final is None:
        saldo_final = sum(item.valor for item in itens)

    linhas: list[str] = [OFX_HEADER.rstrip(), "", "<OFX>"]
    linhas.extend(
        [
            "<SIGNONMSGSRSV1>",
            "<SONRS>",
            "<STATUS>",
            "<CODE>0",
            "<SEVERITY>INFO",
            "</STATUS>",
            f"<DTSERVER>{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "<LANGUAGE>POR",
            "</SONRS>",
            "</SIGNONMSGSRSV1>",
            "<BANKMSGSRSV1>",
            "<STMTTRNRS>",
            "<TRNUID>1001",
            "<STATUS>",
            "<CODE>0",
            "<SEVERITY>INFO",
            "</STATUS>",
            "<STMTRS>",
            f"<CURDEF>{CURDEF}",
            "<BANKACCTFROM>",
            f"<BANKID>{bankid}",
            f"<ACCTID>{acctid}",
            f"<ACCTTYPE>{ACCTTYPE}",
            "</BANKACCTFROM>",
            "<BANKTRANLIST>",
            f"<DTSTART>{dtstart}",
            f"<DTEND>{dtend}",
        ]
    )

    for idx, item in enumerate(itens, start=1):
        linhas.extend(
            [
                "<STMTTRN>",
                f"<TRNTYPE>{item.trntype}",
                f"<DTPOSTED>{item.dtposted}",
                f"<TRNAMT>{formatar_valor_br(item.valor)}",
                f"<FITID>N{idx:05d}",
                f"<CHECKNUM>{item.documento or idx}",
                f"<MEMO>{item.memo}",
                "</STMTTRN>",
            ]
        )

    linhas.extend(
        [
            "</BANKTRANLIST>",
            "<LEDGERBAL>",
            f"<BALAMT>{formatar_valor_br(saldo_final)}",
            "<DTASOF>00000000",
            "</LEDGERBAL>",
            "</STMTRS>",
            "</STMTTRNRS>",
            "</BANKMSGSRSV1>",
            "</OFX>",
        ]
    )
    return "\n".join(linhas) + "\n"


def converter_pdf(pdf_path: Path, out_dir: Path, bankid: str, acctid: str | None) -> Path:
    texto = ler_texto_pdf(pdf_path)
    banco = detectar_banco(texto)
    conta_detectada = extrair_conta(texto, banco)
    conta = acctid if acctid else conta_detectada

    if banco == "evolua":
        lancamentos, saldo_final = parse_evolua(texto)
    else:
        lancamentos, saldo_final = parse_sicredi(texto)

    ofx = montar_ofx(lancamentos, bankid=bankid, acctid=conta, saldo_final=saldo_final)
    out_dir.mkdir(parents=True, exist_ok=True)
    saida = out_dir / f"{pdf_path.stem}.ofx"
    saida.write_text(ofx, encoding="utf-8")
    return saida


def listar_pdfs(caminho: Path) -> list[Path]:
    if caminho.is_file():
        return [caminho]
    return sorted(caminho.glob("*.pdf"))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Converte extratos PDF (Evolua/Sicredi) para OFX no padrao do template."
    )
    parser.add_argument(
        "entrada",
        help="Arquivo PDF ou pasta com PDFs.",
    )
    parser.add_argument(
        "--saida",
        help=f"Pasta de saida dos OFX. Padrao: {PASTA_SAIDA_PADRAO}",
    )
    parser.add_argument("--bankid", default=BANKID_PADRAO, help=f"BANKID no OFX (padrao: {BANKID_PADRAO}).")
    parser.add_argument("--acctid", default=None, help="ACCTID no OFX. Se omitido, tenta detectar do PDF.")
    args = parser.parse_args()

    entrada = Path(args.entrada)
    if not entrada.exists():
        print(f"Erro: caminho nao encontrado: {entrada}")
        return 1

    saida_dir = Path(args.saida) if args.saida else PASTA_SAIDA_PADRAO
    pdfs = listar_pdfs(entrada)
    if not pdfs:
        print("Nenhum PDF encontrado.")
        return 1

    erros = 0
    for pdf in pdfs:
        try:
            ofx_path = converter_pdf(pdf, saida_dir, bankid=args.bankid, acctid=args.acctid)
            print(f"OK: {pdf.name} -> {ofx_path.name}")
        except Exception as exc:
            erros += 1
            print(f"ERRO: {pdf.name}: {exc}")

    return 1 if erros else 0


if __name__ == "__main__":
    raise SystemExit(main())
