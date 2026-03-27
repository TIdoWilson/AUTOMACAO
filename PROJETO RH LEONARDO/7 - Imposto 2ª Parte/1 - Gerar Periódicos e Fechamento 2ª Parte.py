import importlib.util
import sys
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
    if "--skip-darf-fgts" not in sys.argv:
        sys.argv.append("--skip-darf-fgts")

    # Este wrapper e para rodar Periodicos/Fechamento. Pula o pre-flow sempre.
    if "--skip-preflow" not in sys.argv:
        sys.argv.append("--skip-preflow")

    # O script base tenta usar "<nome do arquivo base>.json". Aqui o JSON real e
    # "Imposto 2ª Parte.json", entao forca o caminho se o usuario nao informou.
    if "--coords-file" not in sys.argv:
        base_dir = Path(__file__).resolve().parent
        coords_path = base_dir / "Imposto 2ª Parte.json"
        if coords_path.exists():
            sys.argv.extend(["--coords-file", str(coords_path)])
    mod.main()


if __name__ == "__main__":
    main()
