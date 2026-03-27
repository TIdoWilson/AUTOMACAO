import json
import time
from datetime import datetime
from pathlib import Path

import pyautogui as pag
import uiautomation as uia

BASE_DIR = Path(__file__).resolve().parent
COORDS_PATH = Path(__file__).with_suffix(".json")

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


def type_mes_ano(coord):
    hoje = datetime.now()
    mes_ano = hoje.strftime("%m/%Y")
    click_at(coord)
    time.sleep(WAIT_SHORT)
    pag.press("backspace", presses=7, interval=0.02)
    pag.typewrite(mes_ano, interval=0.02)


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


def main():
    coords = load_coords(COORDS_PATH)

    print("[info] foco em 3s")
    time.sleep(3)

    focus_folha_pagamento()
    time.sleep(WAIT_SHORT)

    click_at((500, 500))
    time.sleep(WAIT_SHORT)

    # ALT + E + F
    alt_sequence("e", "f")
    time.sleep(WAIT_MED)

    # MES/ANO: {mes atual/ano atual}
    type_mes_ano(coords["mes_ano"])
    time.sleep(WAIT_MED)

    # MARCAR "AGRUPADO"
    click_at(coords["marcar_agrupado"])
    time.sleep(WAIT_SHORT)

    # AGRUPAMENTO: 1
    click_at(coords["agrupamento"])
    time.sleep(WAIT_SHORT)
    pag.typewrite("1", interval=0.02)
    pag.press("enter")
    time.sleep(WAIT_SHORT)

    # CONSULTAR
    click_at(coords["consultar"])
    time.sleep(10)

    # POSSUI PAGAMENTO DE REMUNERACAO: MARCAR TODAS (2x)
    menu_action(
        coords["possui_pagamento_remuneracao"],
        keys=["down", "down", "right", "enter"],
        repeats=2,
    )

    # TRANSMISSAO AUTOMATICA DCTFWEB: MARCAR TODAS (2x)
    menu_action(
        coords["transmissao_automatica_dctfweb"],
        keys=["down", "down", "right", "enter"],
        repeats=2,
    )

    # FECHAMENTO: DESMARCAR TODAS (2x)
    menu_action(
        coords["fechamento"],
        keys=["down", "down", "right", "down", "enter"],
        repeats=2,
    )

    # SITUACAO: MARCAR TODAS (2x)
    menu_action(
        coords["situacao"],
        keys=["down", "down", "right", "enter"],
        repeats=2,
    )

    # GERAR
    click_at(coords["gerar"])


if __name__ == "__main__":
    main()
