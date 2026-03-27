import re
from pathlib import Path
from datetime import datetime

# Regex para identificar histórico em branco:
# Linha começando com H, depois somente espaços e no final apenas dígitos (número da linha)
PADRAO_HISTORICO_VAZIO = re.compile(r"^H\s+\d+\s*$")


def processar_arquivo(caminho_arquivo: Path) -> tuple[int, Path | None, Path | None]:
    """
    Processa o arquivo TOSCAN:
    - NÃO sobrescreve o arquivo original.
    - Remove pares L/H cujo histórico está em branco.
    - Cria pasta de resultados ao lado do arquivo (para o monitor mover).
    - Salva arquivo final + arquivo de linhas removidas nessa pasta.
    """

    caminho_arquivo = Path(caminho_arquivo)

    with caminho_arquivo.open("r", encoding="utf-8", errors="replace") as f:
        linhas = f.readlines()

    linhas_mantidas: list[str] = []
    linhas_removidas: list[str] = []

    i = 0
    while i < len(linhas):
        linha_atual = linhas[i]

        # Verifica par L + H
        if linha_atual.startswith("L") and i + 1 < len(linhas):
            proxima_linha = linhas[i + 1]

            if PADRAO_HISTORICO_VAZIO.match(proxima_linha):
                linhas_removidas.append(linha_atual)
                linhas_removidas.append(proxima_linha)
                i += 2
                continue

        linhas_mantidas.append(linha_atual)
        i += 1

    # === CRIA PASTA DE RESULTADOS ===
    pasta_base = caminho_arquivo.parent
    base_prefix = caminho_arquivo.stem.upper()
    pasta_resultado = pasta_base / f"{base_prefix}_RESULTADO"
    pasta_resultado.mkdir(parents=True, exist_ok=True)

    # Arquivo final (linhas mantidas)
    caminho_final = pasta_resultado / caminho_arquivo.name
    with caminho_final.open("w", encoding="utf-8", errors="replace") as f:
        f.writelines(linhas_mantidas)

    # Arquivo de linhas removidas (se houver)
    caminho_removidas = None
    if linhas_removidas:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_removidas = f"linhas_removidas_{caminho_arquivo.stem}_{timestamp}.txt"
        caminho_removidas = pasta_resultado / nome_removidas

        with caminho_removidas.open("w", encoding="utf-8", errors="replace") as f:
            f.writelines(linhas_removidas)

    return len(linhas_removidas), pasta_resultado, caminho_removidas


def main(caminho_arquivo: str | None = None) -> None:
    import sys

    if caminho_arquivo is None:
        if len(sys.argv) < 2:
            print("Uso: python lote_toscan.py <arquivo.txt>")
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
        msg.append("Nenhuma linha com histórico vazio foi encontrada.")

    print("\n".join(msg))


if __name__ == "__main__":
    main()
