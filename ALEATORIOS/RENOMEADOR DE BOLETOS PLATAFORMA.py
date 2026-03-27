import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

# Remove do começo:  "1_Fatura 1-2026_"  (case-insensitive, aceita espaços)
PREFIX_REGEX = re.compile(
    r'^\s*\d+_fatura\s*\d{1,2}-\d{4}_+',
    flags=re.IGNORECASE
)

DRY_RUN = False  # True = só mostra | False = renomeia de verdade


def escolher_pasta() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
    return Path(pasta) if pasta else None


def nome_disponivel(destino: Path) -> Path:
    """Se já existir, cria sufixo (1), (2), ..."""
    if not destino.exists():
        return destino
    base = destino.stem
    ext = destino.suffix
    i = 1
    while True:
        cand = destino.with_name(f"{base} ({i}){ext}")
        if not cand.exists():
            return cand
        i += 1


def main():
    pasta = escolher_pasta()
    if not pasta:
        print("Nenhuma pasta selecionada. Saindo.")
        return

    pdfs = sorted(pasta.glob("*.pdf"))
    print(f"Pasta: {pasta}")
    print(f"Encontrados {len(pdfs)} PDF(s).")
    print("-" * 60)

    alterados = 0
    sem_match = 0

    for arq in pdfs:
        antigo = arq.name
        novo = PREFIX_REGEX.sub("", antigo)

        if novo == antigo:
            sem_match += 1
            continue

        # segurança: evita nome vazio ou só extensão
        if not novo.strip() or novo.strip().lower() == ".pdf":
            print(f"[PULADO] {antigo} -> (resultado inválido)")
            continue

        destino = nome_disponivel(arq.with_name(novo))

        if DRY_RUN:
            print(f"[DRY] {antigo}  ->  {destino.name}")
        else:
            arq.rename(destino)
            print(f"[OK ] {antigo}  ->  {destino.name}")

        alterados += 1

    print("-" * 60)
    if DRY_RUN:
        print(f"DRY_RUN ligado: {alterados} arquivo(s) seriam renomeados.")
        print(f"Sem match (não seguiu o padrão): {sem_match} arquivo(s).")
        print("Para renomear de verdade, mude DRY_RUN = False e rode novamente.")
    else:
        print(f"Concluído: {alterados} arquivo(s) renomeado(s).")
        print(f"Sem match: {sem_match} arquivo(s).")


if __name__ == "__main__":
    main()
