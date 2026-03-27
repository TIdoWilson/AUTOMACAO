import os
from pathlib import Path
from datetime import datetime

PALAVRAS_CHAVE = [
    "rendimento",
    "desconto obtido",
    "pagamento",
    "pagar",
    "adiantamento a fornecedor",
    "adiantamento ao fornecedor",
    "distribuicao",
    "transf. caixa",
    "cesta de relacionamento",
    "tarifa cobranca",
]


def historico_contem_palavra_chave(linha_h: str) -> bool:
    """Verifica se a linha de histórico contém alguma das palavras-chave."""
    texto = linha_h.lower()
    return any(palavra in texto for palavra in PALAVRAS_CHAVE)


def processar_arquivo(caminho_arquivo: Path) -> tuple[int, Path | None, Path | None]:
    """
    Processa o arquivo:
    - NÃO sobrescreve o arquivo original.
    - Gera um arquivo FINAL (linhas mantidas) em uma pasta de resultados.
    - Gera um arquivo de 'linhas removidas' na mesma pasta de resultados.
    - A pasta de resultados é criada ao lado do arquivo, com nome iniciando
      pelo nome do arquivo em maiúsculo, para o monitor conseguir mover
      depois para a pasta 'resultados'.
    """
    caminho_arquivo = Path(caminho_arquivo)

    with caminho_arquivo.open("r", encoding="utf-8", errors="replace") as f:
        linhas = f.readlines()

    linhas_mantidas: list[str] = []
    linhas_removidas: list[str] = []

    i = 0
    while i < len(linhas):
        linha_atual = linhas[i]

        # Se for lançamento e existir uma linha seguinte
        if linha_atual.startswith("L") and i + 1 < len(linhas):
            proxima_linha = linhas[i + 1]

            # Só consideramos histórico se começar com H
            if proxima_linha.startswith("H") and historico_contem_palavra_chave(proxima_linha):
                # Remove L e H (adiciona ambas ao arquivo de removidas)
                linhas_removidas.append(linha_atual)
                linhas_removidas.append(proxima_linha)
                i += 2
                continue

        # Caso não tenha sido removida, mantemos a linha atual
        linhas_mantidas.append(linha_atual)
        i += 1

    # === CRIA PASTA DE RESULTADOS AO LADO DO ARQUIVO (em 'em processamento') ===
    pasta_base = caminho_arquivo.parent                # ...\LOTES INTERNETS\em processamento
    base_prefix = caminho_arquivo.stem.upper()         # nome do arquivo em maiúsculo
    nome_pasta_resultado = f"{base_prefix}_RESULTADO"  # ex.: ARQUIVO_RESULTADO
    pasta_resultado = pasta_base / nome_pasta_resultado
    pasta_resultado.mkdir(parents=True, exist_ok=True)

    # Caminhos dos arquivos de saída
    # arquivo final (linhas mantidas)
    caminho_final = pasta_resultado / caminho_arquivo.name
    # arquivo de linhas removidas
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho_removidas = pasta_resultado / f"linhas_removidas_{caminho_arquivo.stem}_{timestamp}.txt"

    # Grava o arquivo FINAL com as linhas mantidas
    with caminho_final.open("w", encoding="utf-8", errors="replace") as f:
        f.writelines(linhas_mantidas)

    # Grava o arquivo de linhas removidas (se houver)
    if linhas_removidas:
        with caminho_removidas.open("w", encoding="utf-8", errors="replace") as f:
            f.writelines(linhas_removidas)
    else:
        caminho_removidas = None  # nada foi removido

    # NÃO mexemos no arquivo original.
    # O MONITOR irá:
    # - mover o arquivo original de 'em processamento' para 'originais'
    # - mover a pasta *_RESULTADO de 'em processamento' para 'resultados'

    removidas = len(linhas_removidas)
    return removidas, pasta_resultado, caminho_removidas


def main(caminho_arquivo: str | None = None) -> None:
    """
    Modo de uso:
        python "lote_internets.py" "C:\\caminho\\arquivo.txt"

    Quando chamado pelo monitor, ele sempre enviará o caminho do arquivo via argumento.
    """
    import sys

    if caminho_arquivo is None:
        if len(sys.argv) < 2:
            print("Uso: python lote_internets.py <arquivo.txt>")
            return
        caminho_arquivo = sys.argv[1]

    caminho = Path(caminho_arquivo)

    if not caminho.exists():
        print(f"Arquivo não encontrado: {caminho}")
        return

    removidas, pasta_resultado, caminho_removidas = processar_arquivo(caminho)

    msg = [
        "Processamento concluído.",
        f"Pasta de resultados: {pasta_resultado}",
    ]
    if removidas > 0 and caminho_removidas:
        msg.append(f"Linhas removidas (L + H): {removidas}")
        msg.append(f"Arquivo de linhas removidas: {caminho_removidas}")
    else:
        msg.append("Nenhum histórico com as palavras-chave foi encontrado.")

    print("\n".join(msg))


if __name__ == "__main__":
    main()
