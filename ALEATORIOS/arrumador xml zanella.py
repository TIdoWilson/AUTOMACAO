import re
from pathlib import Path
from tkinter import Tk, filedialog, messagebox


XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8"?>'
NFE_NAMESPACE = "http://www.portalfiscal.inf.br/nfe"
PROTNFE_VERSAO = "4.00"
PROTNFE_INFPROT_ID = "ID141260352378914"
PROTNFE_TPAMB = "1"
PROTNFE_VERAPLIC = "PR-v4_5_29"
PROTNFE_DHRECBTO = "2026-03-02T08:24:18-03:00"
PROTNFE_NPROT = "141260352378914"
PROTNFE_DIGVAL = "2d4DJf39KLpSEI7wz1PSad4nOgU="
PROTNFE_CSTAT = "100"
PROTNFE_XMOTIVO = "Autorizado o uso da NF-e"


def remover_declaracao_xml(texto: str) -> str:
    return re.sub(r"^\s*<\?xml[^>]*\?>\s*", "", texto, count=1, flags=re.IGNORECASE)


def extrair_chave_nfe(texto_xml: str) -> str:
    match = re.search(r'Id\s*=\s*"NFe(\d{44})"', texto_xml)
    if not match:
        raise ValueError("Nao foi possivel encontrar a chave da NFe no atributo Id de infNFe.")
    return match.group(1)


def gerar_bloco_protnfe(chave: str) -> str:
    return (
        f'<protNFe xmlns="{NFE_NAMESPACE}" versao="{PROTNFE_VERSAO}">\n'
        f'    <infProt Id="{PROTNFE_INFPROT_ID}">\n'
        f"      <tpAmb>{PROTNFE_TPAMB}</tpAmb>\n"
        f"      <verAplic>{PROTNFE_VERAPLIC}</verAplic>\n"
        f"      <chNFe>{chave}</chNFe>\n"
        f"      <dhRecbto>{PROTNFE_DHRECBTO}</dhRecbto>\n"
        f"      <nProt>{PROTNFE_NPROT}</nProt>\n"
        f"      <digVal>{PROTNFE_DIGVAL}</digVal>\n"
        f"      <cStat>{PROTNFE_CSTAT}</cStat>\n"
        f"      <xMotivo>{PROTNFE_XMOTIVO}</xMotivo>\n"
        f"    </infProt>\n"
        f"  </protNFe>"
    )


def montar_xml_ajustado(texto_xml: str) -> str:
    texto_limpo = remover_declaracao_xml(texto_xml).strip()
    chave = extrair_chave_nfe(texto_limpo)
    bloco_protnfe = gerar_bloco_protnfe(chave)

    if re.search(r"<nfeProc\b", texto_limpo, flags=re.IGNORECASE):
        sem_protnfe = re.sub(
            r"\s*<protNFe\b[\s\S]*?</protNFe>\s*",
            "\n",
            texto_limpo,
            count=1,
            flags=re.IGNORECASE,
        )

        if not re.search(r"</nfeProc>\s*$", sem_protnfe, flags=re.IGNORECASE):
            raise ValueError("Arquivo com nfeProc invalido: fechamento </nfeProc> nao encontrado.")

        corpo = re.sub(
            r"</nfeProc>\s*$",
            f"\n  {bloco_protnfe}\n</nfeProc>",
            sem_protnfe,
            count=1,
            flags=re.IGNORECASE,
        )
        return XML_DECLARATION + "\n" + corpo.strip() + "\n"

    corpo = (
        f'<nfeProc xmlns="{NFE_NAMESPACE}" versao="4.00">\n'
        f"  {texto_limpo}\n"
        f"  {bloco_protnfe}\n"
        f"</nfeProc>"
    )
    return XML_DECLARATION + "\n" + corpo + "\n"


def nome_saida(nome_original: str) -> str:
    if nome_original.lower().startswith("alterado"):
        return f"ajustado_{nome_original}"
    return f"alterado{nome_original}"


def executar() -> None:
    root = Tk()
    root.withdraw()

    modo_pasta = messagebox.askyesno(
        "Ajustador XML NFC-e",
        "Deseja selecionar uma pasta para processar todos os XMLs dela?\n\nSim: escolhe pasta (processa lote inteiro).\nNao: escolhe arquivos manualmente.",
    )

    if modo_pasta:
        pasta_alvo = filedialog.askdirectory(title="Selecione a pasta com os XMLs para ajustar")
        if not pasta_alvo:
            messagebox.showinfo("Ajustador XML NFC-e", "Operacao cancelada: pasta nao selecionada.")
            return

        caminhos_alvo = tuple(str(p) for p in Path(pasta_alvo).glob("*.xml"))
        if not caminhos_alvo:
            messagebox.showinfo("Ajustador XML NFC-e", "Nenhum arquivo .xml encontrado na pasta selecionada.")
            return
    else:
        caminhos_alvo = filedialog.askopenfilenames(
            title="Selecione os XMLs para ajustar (use Ctrl/Shift para varios)",
            filetypes=[("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")],
        )
        if not caminhos_alvo:
            messagebox.showinfo("Ajustador XML NFC-e", "Operacao cancelada: nenhum XML alvo selecionado.")
            return

    sobrescrever = messagebox.askyesno(
        "Ajustador XML NFC-e",
        "Deseja sobrescrever os arquivos originais?\n\nSim: sobrescreve os XMLs selecionados.\nNao: cria novos arquivos com prefixo alterado.",
    )

    pasta_saida = ""
    if not sobrescrever:
        pasta_saida = filedialog.askdirectory(
            title="Selecione a pasta de saida (Cancelar = mesma pasta de cada XML)"
        )

    ok = 0
    erros: list[str] = []

    for caminho in caminhos_alvo:
        try:
            origem = Path(caminho)
            texto = origem.read_text(encoding="utf-8", errors="replace")
            ajustado = montar_xml_ajustado(texto)

            if sobrescrever:
                destino = origem
            else:
                base_dir = Path(pasta_saida) if pasta_saida else origem.parent
                destino = base_dir / nome_saida(origem.name)

            destino.write_text(ajustado, encoding="utf-8", newline="\n")
            ok += 1
        except Exception as exc:  # noqa: BLE001
            erros.append(f"{Path(caminho).name}: {exc}")

    resumo = [f"Arquivos processados com sucesso: {ok}"]
    if erros:
        resumo.append(f"Arquivos com erro: {len(erros)}")
        resumo.append("")
        resumo.append("Detalhes:")
        resumo.extend(erros)

    messagebox.showinfo("Ajustador XML NFC-e", "\n".join(resumo))


if __name__ == "__main__":
    executar()
