import os
from pathlib import Path


def load_local_env(*paths: str) -> None:
    for raw_path in paths:
        if not raw_path:
            continue

        path = Path(raw_path)
        if not path.is_file():
            continue

        try:
            with path.open("r", encoding="utf-8") as handle:
                for line in handle:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue

                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()

                    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
                        value = value[1:-1]

                    if key and key not in os.environ:
                        os.environ[key] = value
        except OSError:
            continue
