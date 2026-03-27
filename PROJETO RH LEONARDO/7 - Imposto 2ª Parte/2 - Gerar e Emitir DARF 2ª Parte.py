import argparse
import importlib.util
from pathlib import Path

SOURCE_SCRIPT = "backup_Imposto 2ª Parte.py"


def _load_imposto_module():
    base_dir = Path(__file__).resolve().parent
    src_path = base_dir / SOURCE_SCRIPT
    if not src_path.exists():
        raise FileNotFoundError(f"Arquivo base nao encontrado: {src_path}")
    spec = importlib.util.spec_from_file_location("imposto_2a_parte", src_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Falha ao carregar modulo: {src_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main():
    mod = _load_imposto_module()

    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--darf-coords-file",
        default=str(mod.DARF_COORDS_PATH),
        help="Arquivo de coordenadas do DARF.",
    )
    ap.add_argument(
        "--darf-grupos",
        default="31;32;33;34;35;36;37",
        help="Lista de grupos DARF (ex: 31;32;33;34;35;36;37).",
    )
    ap.add_argument(
        "--skip-emitir-darf",
        action="store_true",
        help="Nao clica em Emitir DARF.",
    )
    ap.add_argument(
        "--skip-organizar-darf",
        action="store_true",
        help="Nao organiza/renomeia PDFs de DARF na pasta automatizado.",
    )
    args = ap.parse_args()

    darf_grupos = mod.parse_grupos(args.darf_grupos or "")
    if not darf_grupos:
        darf_input = input("Grupos DARF (ex: 6;13;14): ").strip()
        darf_grupos = mod.parse_grupos(darf_input)
    if not darf_grupos:
        raise SystemExit("Grupos DARF nao informados.")

    darf_coords = mod.load_coords_darf(Path(args.darf_coords_file))

    mod.run_darf_flow(
        darf_coords,
        grupos=darf_grupos,
        skip_emitir=args.skip_emitir_darf,
        skip_organizar=args.skip_organizar_darf,
    )


if __name__ == "__main__":
    main()

