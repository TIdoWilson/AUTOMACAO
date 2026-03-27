import os
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path
from shutil import which

PORT = 9222


def find_chrome_exe() -> str:
    candidates = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
        which("chrome"),
        which("chrome.exe"),
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    raise FileNotFoundError("chrome.exe nao encontrado. Ajuste o caminho manualmente.")


def cdp_ready(port: int = PORT) -> bool:
    url = f"http://127.0.0.1:{port}/json/version"
    try:
        with urllib.request.urlopen(url, timeout=1) as r:
            return r.status == 200
    except Exception:
        return False


def launch_chrome_with_cdp(port: int = PORT):
    chrome = find_chrome_exe()
    user_data = str(Path(tempfile.gettempdir()) / f"chrome-dev-{port}")
    os.makedirs(user_data, exist_ok=True)
    args = [
        chrome,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={user_data}",
        "--incognito",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(60):  # ~15s
        if cdp_ready(port):
            return proc
        time.sleep(0.25)
    raise TimeoutError("Chrome com CDP nao respondeu a tempo.")


def ensure_chrome_cdp(port: int = PORT):
    if not cdp_ready(port):
        launch_chrome_with_cdp(port)
