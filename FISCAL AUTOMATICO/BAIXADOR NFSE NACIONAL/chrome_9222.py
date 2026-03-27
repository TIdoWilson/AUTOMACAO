import os
import time
import tempfile
import subprocess
import urllib.request
from pathlib import Path
from shutil import which
from playwright.sync_api import sync_playwright

PORT = 9222

def find_chrome_exe():
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
    raise FileNotFoundError("chrome.exe não encontrado. Ajuste o caminho manualmente.")

def cdp_ready(port=PORT):
    url = f"http://127.0.0.1:{port}/json/version"
    try:
        with urllib.request.urlopen(url, timeout=1) as r:
            return r.status == 200
    except Exception:
        return False

def launch_chrome_with_cdp(port=PORT):
    chrome = find_chrome_exe()
    user_data = str(Path(tempfile.gettempdir()) / f"chrome-dev-{port}")
    os.makedirs(user_data, exist_ok=True)
    args = [
        chrome,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={user_data}",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(60):  # ~15s
        if cdp_ready(port):
            return proc
        time.sleep(0.25)
    raise TimeoutError("Chrome com CDP não respondeu a tempo.")

def chrome_9222(p, port=PORT):
    """Retorna browser conectado ao Chrome via CDP."""
    if not cdp_ready(port):
        launch_chrome_with_cdp(port)
    return p.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
