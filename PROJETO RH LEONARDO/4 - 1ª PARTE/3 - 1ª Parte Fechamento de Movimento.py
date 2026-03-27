import argparse
import json
import time
from datetime import datetime
from pathlib import Path

import pyautogui as pag
import uiautomation as uia

COORDS_PATH = Path(__file__).with_suffix(".json")

DEFAULT_AGRUPAMENTO = "6;13;14"

WAIT_SHORT = 0.2
WAIT_MED = 0.8

pag.PAUSE = 0.05
pag.FAILSAFE = False


def load_coords(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de coordenadas nao encontrado: {path}")
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    required = [
        "mes_ano",
        "marcar_agrupado",
        "agrupamento",
        "consultar",
        "possui_remuneracao",
        "possui_pagamento_remuneracao",
        "transmissao_automatica_dctfweb",
        "fechamento",
        "situacao",
        "gerar",
    ]
    for key in required:
        if key not in data:
            raise KeyError(f"Coordenada ausente no JSON: {key}")
        coord = data[key]
        if (
            not isinstance(coord, (list, tuple))
            or len(coord) != 2
            or not all(isinstance(v, (int, float)) for v in coord)
        ):
            raise ValueError(f"Coordenada invalida para {key}: {coord}")
    return data


def click_at(coord):
    x, y = coord
    pag.moveTo(x, y, duration=0.1)
    pag.click()


def right_click_at(coord):
    x, y = coord
    pag.moveTo(x, y, duration=0.1)
    pag.click(button="right")


def type_text(coord, text: str):
    click_at(coord)
    time.sleep(WAIT_SHORT)
    pag.press("backspace", presses=8, interval=0.02)
    pag.press("delete", presses=8, interval=0.02)
    pag.typewrite(text, interval=0.02)


def type_mes_ano(coord):
    hoje = datetime.now()
    mes_ano = hoje.strftime("%m/%Y")
    type_text(coord, mes_ano)
    pag.press("enter")


def alt_sequence(*keys):
    pag.keyDown("alt")
    for k in keys:
        pag.press(k)
        time.sleep(0.05)
    pag.keyUp("alt")


def menu_action(coord, keys, repeats=2):
    for _ in range(repeats):
        right_click_at(coord)
        time.sleep(0.5)
        for k in keys:
            pag.press(k)
            time.sleep(0.5)


def focus_folha_pagamento():
    wnd = uia.WindowControl(Name="Folha de Pagamento", ClassName="TfrmPrincipal")
    if not wnd.Exists(2):
        return False
    try:
        wnd.SetFocus()
        return True
    except Exception:
        return False


def run_flow(coords, agrupamento: str, wait_consulta: float, wait_gerar: float):
    # ALT + E + F
    alt_sequence("e", "f")
    time.sleep(WAIT_MED)

    # MES/ANO
    type_mes_ano(coords["mes_ano"])
    time.sleep(WAIT_MED)

    # AGRUPADO + AGRUPAMENTO
    click_at(coords["marcar_agrupado"])
    time.sleep(WAIT_SHORT)
    type_text(coords["agrupamento"], str(agrupamento))
    pag.press("enter")
    time.sleep(WAIT_SHORT)

    # CONSULTAR
    click_at(coords["consultar"])
    time.sleep(wait_consulta)

    # COLUNAS: MARCAR TODAS (2x)
    menu_action(coords["possui_remuneracao"], keys=["down", "down", "right", "enter"], repeats=2)
    menu_action(
        coords["possui_pagamento_remuneracao"], keys=["down", "down", "right", "enter"], repeats=2
    )
    menu_action(
        coords["transmissao_automatica_dctfweb"], keys=["down", "down", "right", "enter"], repeats=2
    )
    menu_action(coords["situacao"], keys=["down", "down", "right", "enter"], repeats=2)

    # FECHAMENTO: DESMARCAR TODAS (2x)
    menu_action(
        coords["fechamento"],
        keys=["down", "down", "right", "down", "enter"],
        repeats=2,
    )

    # GERAR
    click_at(coords["gerar"])
    time.sleep(wait_gerar)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--coords-file",
        default=None,
        help="Arquivo de coordenadas (.json). Padrao: arquivo com o mesmo nome do script.",
    )
    ap.add_argument(
        "--agrupamento",
        default=DEFAULT_AGRUPAMENTO,
        help="Texto do agrupamento (padrao: 6;13;14).",
    )
    ap.add_argument("--wait-consulta", type=float, default=10.0, help="Tempo de espera apos consultar.")
    ap.add_argument("--wait-gerar", type=float, default=10.0, help="Tempo de espera apos gerar.")
    args = ap.parse_args()

    coords_path = Path(args.coords_file) if args.coords_file else COORDS_PATH
    coords = load_coords(coords_path)

    print("[info] foco em 3s")
    time.sleep(3)

    focus_folha_pagamento()
    time.sleep(WAIT_SHORT)

    run_flow(coords, args.agrupamento, args.wait_consulta, args.wait_gerar)


if __name__ == "__main__":
    main()
