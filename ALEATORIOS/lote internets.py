import os
import tkinter as tk
from tkinter import filedialog, messagebox

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
    "tarifa cobranca"
]

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo de lançamentos",
        filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")]
    )
    root.destroy()
    return caminho

def historico_contem_palavra_chave(linha_h):
    """Verifica se a linha de histórico contém alguma das palavras-chave."""
    texto = linha_h.lower()
    return any(palavra in texto for palavra in PALAVRAS_CHAVE)

def processar_arquivo(caminho_arquivo):
    # Lê todas as linhas do arquivo
    with open(caminho_arquivo, "r", encoding="utf-8", errors="replace") as f:
        linhas = f.readlines()

    linhas_mantidas = []
    linhas_removidas = []

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

    # Regrava o arquivo original somente com as linhas mantidas
    with open(caminho_arquivo, "w", encoding="utf-8", errors="replace") as f:
        f.writelines(linhas_mantidas)

    # Cria o arquivo "linhas removidas.txt" na mesma pasta, se houver algo removido
    if linhas_removidas:
        pasta = os.path.dirname(caminho_arquivo)
        caminho_removidas = os.path.join(pasta, "linhas removidas.txt")

        with open(caminho_removidas, "w", encoding="utf-8", errors="replace") as f:
            f.writelines(linhas_removidas)

        return len(linhas_removidas), caminho_removidas
    else:
        return 0, None

def main():
    caminho_arquivo = selecionar_arquivo()
    if not caminho_arquivo:
        print("Nenhum arquivo selecionado. Encerrando.")
        return

    removidas, caminho_removidas = processar_arquivo(caminho_arquivo)

    root = tk.Tk()
    root.withdraw()

    if removidas > 0:
        messagebox.showinfo(
            "Concluído",
            f"Processamento concluído.\n"
            f"Linhas removidas (L + H): {removidas}\n"
            f"Arquivo criado:\n{caminho_removidas}"
        )
    else:
        messagebox.showinfo(
            "Concluído",
            "Nenhum histórico com as palavras-chave foi encontrado."
        )

    root.destroy()

if __name__ == "__main__":
    main()
