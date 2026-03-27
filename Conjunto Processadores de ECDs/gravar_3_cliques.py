# -*- coding: utf-8 -*-
"""
Grava 3 cliques (posicao do mouse) ao apertar ENTER.
"""

from __future__ import annotations

try:
    from pynput import mouse
except Exception as exc:
    raise SystemExit("pynput é obrigatório para capturar cliques") from exc


def main() -> None:
    coords = []
    max_clicks = 20
    print("Clique 20 vezes para registrar as coordenadas.")

    def on_click(x, y, button, pressed):
        if not pressed:
            return
        coords.append((x, y))
        print(f"Capturado {len(coords)}/{max_clicks}: ({x}, {y})")
        if len(coords) >= max_clicks:
            return False

    with mouse.Listener(on_click=on_click) as listener:
        listener.join()

    out_path = r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\Conjunto Processadores de ECDs\coordenadas entrada DFC.txt"

    linhas = [f"{i+1}: ({c[0]}, {c[1]})" for i, c in enumerate(coords)]

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas) + "\n")

    print("Coordenadas:")
    for ln in linhas:
        print(ln)
    print(f"Salvo em: {out_path}")


if __name__ == "__main__":
    main()
