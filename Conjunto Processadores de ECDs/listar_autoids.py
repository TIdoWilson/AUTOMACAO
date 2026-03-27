# -*- coding: utf-8 -*-
"""
Lista AutoID/Name/Class de controles visiveis (win32 + uia).
"""
import os
import time
from pywinauto import Desktop


IGNORAR_JANELAS = {
    "Program Manager",
    "Gerenciador de sistemas",
    "Windows PowerShell",
    "PowerShell",
    "Firefox",
    "Firefox Private Browsing",
    "Firefox - Navegação Privada",
    "Window Spy for AHKv2",
}

def _deve_ignorar_janela(titulo: str, class_name: str) -> bool:
    if class_name == "Shell_TrayWnd":
        return True
    t = (titulo or "").strip()
    if t in IGNORAR_JANELAS:
        return True
    if "Firefox" in t:
        return True
    if "Visual Studio Code" in t:
        return True
    return False


def listar_controles(backend: str, out) -> None:
    print(f"\n=== BACKEND: {backend} ===")
    out.write(f"\n=== BACKEND: {backend} ===\n")
    desk = Desktop(backend=backend)
    for w in desk.windows():
        try:
            titulo = w.window_text()
            cls_win = w.class_name()
            if _deve_ignorar_janela(titulo, cls_win):
                continue
            linha_janela = f"\n[Janela] {titulo} | class={cls_win} | handle={w.handle}"
            print(linha_janela)
            out.write(linha_janela + "\n")
            for c in w.descendants():
                try:
                    auto_id = getattr(c.element_info, "automation_id", "")
                    name = c.window_text()
                    cls = c.class_name()
                    if auto_id == "Hotkey":
                        continue
                    if auto_id or name:
                        linha = f"  auto_id='{auto_id}' | name='{name}' | class='{cls}' | handle={c.handle}"
                        print(linha)
                        out.write(linha + "\n")
                except Exception:
                    continue
        except Exception:
            continue


def _capturar_snapshot_win32() -> set[tuple[str, str, str, str, str]]:
    snapshot: set[tuple[str, str, str, str, str]] = set()
    desk = Desktop(backend="win32")
    for w in desk.windows():
        try:
            titulo = w.window_text()
            cls_win = w.class_name()
            if _deve_ignorar_janela(titulo, cls_win):
                continue
            for c in w.descendants():
                try:
                    auto_id = getattr(c.element_info, "automation_id", "")
                    name = c.window_text()
                    cls = c.class_name()
                    if auto_id == "Hotkey":
                        continue
                    if auto_id or name:
                        snapshot.add((cls_win, titulo, auto_id, name, cls))
                except Exception:
                    continue
        except Exception:
            continue
    return snapshot


if __name__ == "__main__":
    out_path = os.path.join(os.path.dirname(__file__), "listar_autoids_output.txt")
    with open(out_path, "w", encoding="utf-8") as out:
        print("Captura inicial (baseline)...")
        listar_controles("win32", out)
        listar_controles("uia", out)
        before = _capturar_snapshot_win32()
        print("\nAbra o relatório e aguarde 10 segundos...")
        time.sleep(10)
        out.write("\n\n=== DIFERENCAS (DEPOIS - ANTES) ===\n")
        after = _capturar_snapshot_win32()
        novos = sorted(after - before)
        for cls_win, titulo, auto_id, name, cls in novos:
            linha = f"[NOVO] janela='{titulo}' class='{cls_win}' auto_id='{auto_id}' name='{name}' class_ctrl='{cls}'"
            print(linha)
            out.write(linha + "\n")
    print(f"\nArquivo gerado: {out_path}")
