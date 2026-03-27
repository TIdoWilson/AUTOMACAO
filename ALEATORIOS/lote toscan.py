import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# Regex para identificar histórico em branco:
# Linha começando com H, depois somente espaços e no final apenas dígitos (número da linha)
PADRAO_HISTORICO_VAZIO = re.compile(r"^H\s+\d+\s*$")

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo de lançamentos",
        filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")]
    )
    root.destroy()
    return caminho

def processar_arquivo(caminho_arquivo):
    # Lê todas as linhas do arquivo
    with open(caminho_arquivo, "r", encoding="utf-8", errors="replace") as f:
        linhas = f.readlines()

    linhas_mantidas = []
    linhas_removidas = []

    i = 0
    while i < len(linhas):
        linha_atual = linhas[i]

        # Verifica se é uma linha de lançamento (L) e se há uma próxima linha (H) de histórico vazio
        if linha_atual.startswith("L") and i + 1 < len(linhas):
            proxima_linha = linhas[i + 1]

            if PADRAO_HISTORICO_VAZIO.match(proxima_linha):
                # Remove L e H (adiciona ambas ao arquivo de removidas)
                linhas_removidas.append(linha_atual)
                linhas_removidas.append(proxima_linha)
                i += 2
                continue  # pula para o próximo par

        # Se não entrou no caso de remoção, mantemos a linha atual
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
            f"Linhas removidas: {removidas}\n"
            f"Arquivo criado:\n{caminho_removidas}"
        )
    else:
        messagebox.showinfo(
            "Concluído",
            "Nenhuma linha com histórico em branco foi encontrada."
        )

    root.destroy()

if __name__ == "__main__":
    main()
