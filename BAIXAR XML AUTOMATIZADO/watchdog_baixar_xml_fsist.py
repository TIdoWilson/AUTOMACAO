from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Optional

try:
    import tkinter as tk
    from tkinter import filedialog
except Exception:  # pragma: no cover
    tk = None
    filedialog = None


RESUME_DELAY_SECONDS = 3
def pick_file(title: str) -> Optional[Path]:
    if tk is None or filedialog is None:
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
    )
    root.destroy()
    return Path(selected) if selected else None


def pick_folder(title: str, initial_dir: Optional[Path] = None) -> Optional[Path]:
    if tk is None or filedialog is None:
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askdirectory(
        title=title,
        initialdir=str(initial_dir) if initial_dir else None,
        mustexist=False,
    )
    root.destroy()
    return Path(selected) if selected else None


def state_path_for_excel(excel_path: Path) -> Path:
    digest = hashlib.md5(str(excel_path.resolve()).encode("utf-8"), usedforsecurity=False).hexdigest()
    return Path(tempfile.gettempdir()) / f"baixar_xml_fsist_watchdog_{digest}.json"


def resolve_excel_path(arg_path: Optional[Path]) -> Path:
    if arg_path is not None:
        return arg_path
    if tk is None or filedialog is None:
        raise SystemExit("Tkinter indisponivel. Informe o arquivo com --excel.")
    selected = pick_file("Selecione o Excel com as chaves")
    if not selected:
        raise SystemExit("Operacao cancelada: arquivo Excel nao selecionado.")
    return selected


def resolve_output_dir(arg_path: Optional[Path], excel_path: Path) -> Path:
    if arg_path is not None:
        return arg_path
    if tk is not None and filedialog is not None:
        selected = pick_folder("Selecione a pasta para salvar os XMLs", excel_path.parent)
        if selected:
            return selected
    return excel_path.parent / f"{excel_path.stem}_XML"


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Watchdog para retomar o download de XMLs da NF-e.")
    parser.add_argument("--excel", type=Path, help="Arquivo Excel com as chaves.")
    parser.add_argument("--saida", type=Path, help="Pasta para salvar os XMLs.")
    return parser.parse_args(argv)


def run_downloader(script_path: Path, excel_path: Path, output_dir: Path, resume_state: Optional[Path]) -> int:
    cmd = [sys.executable, str(script_path), "--excel", str(excel_path), "--saida", str(output_dir)]
    if resume_state is not None:
        cmd.extend(["--resume-state", str(resume_state)])
    completed = subprocess.run(cmd)
    return completed.returncode


def load_state(state_path: Path) -> dict:
    return json.loads(state_path.read_text(encoding="utf-8"))


def press_space_globally() -> None:
    if not sys.platform.startswith("win"):
        return
    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            "$sig='[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(System.IntPtr hWnd);'; "
            "Add-Type -MemberDefinition $sig -Name NativeMethods -Namespace Win32 -ErrorAction SilentlyContinue | Out-Null; "
            "$chrome = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1; "
            "if ($chrome) { [Win32.NativeMethods]::SetForegroundWindow($chrome.MainWindowHandle) | Out-Null; Start-Sleep -Milliseconds 150 }; "
            "$wsh = New-Object -ComObject WScript.Shell; Start-Sleep -Milliseconds 150; $wsh.SendKeys(' ')",
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def main() -> None:
    args = parse_args(sys.argv[1:])
    excel_path = resolve_excel_path(args.excel)
    if not excel_path.exists():
        raise SystemExit(f"Arquivo Excel nao encontrado: {excel_path}")

    output_dir = resolve_output_dir(args.saida, excel_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    state_path = state_path_for_excel(excel_path)
    downloader_script = Path(__file__).resolve().with_name("baixar_xml_fsist.py")
    if not downloader_script.exists():
        raise SystemExit(f"Downloader nao encontrado: {downloader_script}")

    if state_path.exists():
        state_path.unlink()

    resume_state = None
    while True:
        code = run_downloader(downloader_script, excel_path, output_dir, resume_state)
        if code != 0:
            raise SystemExit(code)
        if not state_path.exists():
            raise SystemExit(code)

        state = load_state(state_path)
        if str(state.get("stage", "")) == "consult":
            press_space_globally()

        time.sleep(RESUME_DELAY_SECONDS)
        resume_state = state_path


if __name__ == "__main__":
    main()
