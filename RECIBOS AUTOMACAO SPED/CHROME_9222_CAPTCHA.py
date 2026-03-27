import os
import time
import subprocess
import urllib.request
from pathlib import Path
from shutil import which

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

def _default_profile_dir(port=PORT):
    # perfil persistente ao lado deste arquivo (ex.: ...\chrome_profile_9222)
    base = Path(__file__).resolve().parent / f"chrome_profile_{port}"
    base.mkdir(parents=True, exist_ok=True)
    return str(base)

def launch_chrome_with_cdp(port=PORT, user_data_dir=None, profile_name=None):
    chrome = find_chrome_exe()
    user_data = user_data_dir or _default_profile_dir(port)

    args = [
        chrome,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={user_data}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-features=BlockThirdPartyCookies",  # permite 3p cookies
        "--disable-extensions",                      # evita bloqueios por extensão
        "--start-maximized",
    ]
    if profile_name:
        args.append(f"--profile-directory={profile_name}")

    proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(120):
        if cdp_ready(port):
            return proc
        time.sleep(0.25)
    raise TimeoutError("Chrome com CDP não respondeu a tempo.")


def get_browser(p, port=PORT, user_data_dir=None, profile_name=None):
    """
    p: async_playwright() ou sync_playwright()
    """
    if not cdp_ready(port):
        launch_chrome_with_cdp(port, user_data_dir=user_data_dir, profile_name=profile_name)
    return p.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
