import argparse
import json
import time
from pathlib import Path

import pyautogui as pag
import uiautomation as uia

WAIT_SHORT = 0.3
WAIT_MED = 0.8
WAIT_LONG = 2.0
WAIT_FIVE_MIN = 300.0
WAIT_BUTTON_TIMEOUT = 60.0

pag.PAUSE = 0.05
pag.FAILSAFE = False


COORDS_PATH = Path(__file__).with_suffix(".json")


def load_coords(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de coordenadas nao encontrado: {path}")
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    required = [
        "modulos",
        "emprestimo_credito_trabalhador",
        "importacao_contratos",
        "consulta_portal_emprega_brasil",
        "empresa_1_1194",
        "consultar",
        "marcar_todas",
        "importar_dados",
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


def click_at(coord, wait_s=WAIT_SHORT):
    x, y = coord
    pag.moveTo(x, y, duration=0.1)
    pag.click()
    time.sleep(wait_s)


def move_to(coord, wait_s=WAIT_SHORT):
    x, y = coord
    pag.moveTo(x, y, duration=0.1)
    time.sleep(wait_s)


def focus_folha_pagamento():
    wnd = uia.WindowControl(Name="Folha de Pagamento", ClassName="TfrmPrincipal")
    if not wnd.Exists(2):
        return False
    try:
        wnd.SetFocus()
        return True
    except Exception:
        return False


def wait_for_button_enabled(names, timeout_s: float = WAIT_BUTTON_TIMEOUT) -> bool:
    end = time.time() + timeout_s
    while time.time() < end:
        wnd = uia.WindowControl(Name="Folha de Pagamento", ClassName="TfrmPrincipal")
        if wnd.Exists(1):
            for name in names:
                try:
                    btn = wnd.ButtonControl(Name=name)
                    if btn.Exists(0.5) and btn.IsEnabled:
                        return True
                except Exception:
                    continue
        time.sleep(0.3)
    return False


def main(wait_enabled: bool, wait_5_min: bool, test_no_import: bool):
    coords = load_coords(COORDS_PATH)
    print("[info] foco em 3s")
    time.sleep(3)

    focus_folha_pagamento()
    time.sleep(WAIT_SHORT)

    click_at(coords["modulos"])
    click_at(coords["emprestimo_credito_trabalhador"])
    click_at(coords["importacao_contratos"], wait_s=WAIT_MED)

    click_at(coords["consulta_portal_emprega_brasil"])
    click_at(coords["empresa_1_1194"])
    empresa_texto = (coords.get("empresa_texto") or "").strip()
    if empresa_texto:
        pag.hotkey("ctrl", "a")
        time.sleep(WAIT_SHORT)
        pag.typewrite(empresa_texto, interval=0.02)
    click_at(coords["consultar"], wait_s=WAIT_LONG)

    if wait_5_min:
        time.sleep(WAIT_FIVE_MIN)
    elif wait_enabled and not wait_for_button_enabled(["Marcar todas", "Marcar Todas"]):
        time.sleep(WAIT_FIVE_MIN)
    click_at(coords["marcar_todas"])

    if test_no_import:
        move_to(coords["importar_dados"], wait_s=WAIT_SHORT)
        print("[teste] mouse posicionado em importar dados (sem clique).")
        return

    if wait_5_min:
        time.sleep(WAIT_FIVE_MIN)
    elif wait_enabled and not wait_for_button_enabled(["Importar dados", "Importar Dados"]):
        time.sleep(WAIT_FIVE_MIN)
    click_at(coords["importar_dados"])


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--wait-enabled",
        action="store_true",
        help="Espera o botao ficar habilitado antes de clicar.",
    )
    ap.add_argument(
        "--wait-5-min",
        action="store_true",
        help="Espera 5 minutos antes de clicar (ignora deteccao).",
    )
    ap.add_argument(
        "--test-no-import",
        action="store_true",
        help="Nao clica em Importar Dados; apenas move o mouse.",
    )
    args = ap.parse_args()
    main(
        wait_enabled=args.wait_enabled,
        wait_5_min=args.wait_5_min,
        test_no_import=args.test_no_import,
    )
