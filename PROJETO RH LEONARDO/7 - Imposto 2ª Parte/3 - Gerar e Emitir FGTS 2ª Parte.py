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
        "--fgts-coords-file",
        default=str(mod.FGTS_COORDS_PATH),
        help="Arquivo de coordenadas do FGTS.",
    )
    ap.add_argument(
        "--fgts-grupos",
        default="31;32;33;34;35;36",
        help="Lista de grupos FGTS (ex: 31;32;33;34;35;36).",
    )
    ap.add_argument(
        "--skip-emitir-fgts",
        action="store_true",
        help="Nao clica em Emitir FGTS.",
    )
    args = ap.parse_args()

    fgts_grupos = mod.parse_grupos(args.fgts_grupos or "")
    if not fgts_grupos:
        fgts_input = input("Grupos FGTS (ex: 6;13;14): ").strip()
        fgts_grupos = mod.parse_grupos(fgts_input)
    if not fgts_grupos:
        raise SystemExit("Grupos FGTS nao informados.")

    fgts_coords = mod.load_coords_fgts(Path(args.fgts_coords_file))

    mod.run_fgts_flow(
        fgts_coords,
        grupos=fgts_grupos,
        skip_emitir=args.skip_emitir_fgts,
    )


if __name__ == "__main__":
    main()
