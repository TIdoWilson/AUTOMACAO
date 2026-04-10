#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import csv
import ctypes
import os
import re
import threading
import time
from pathlib import Path
from tkinter import Tk, filedialog, messagebox, StringVar
from xml.etree import ElementTree as ET

if os.name != "nt":
    raise SystemExit("Este script foi feito para Windows.")


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

VK_F8 = 0x77
VK_CONTROL = 0x11
VK_V = 0x56
KEYEVENTF_KEYUP = 0x0002
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002

PAUSA_INICIAL = 3.0
PAUSA_DEPOIS_CAMPO = 0.60
PAUSA_DEPOIS_COLAR = 0.90
PAUSA_ENTRE_BLOCO = 1.20

COORD_FIELD = (685, 304)
COORD_1 = (1050, 349)
COORD_2 = (272, 435)
COORD_3 = (1096, 417)
COORD_4 = (763, 346)

EXTENSOES_SUPORTADAS = {".xlsx", ".xlsm", ".csv"}
NOMES_DE_COLUNA = {
    "chave",
    "chave de acesso",
    "chave acesso",
    "chave da nota",
    "chave da nfe",
    "chave nfe",
    "chave nf-e",
    "chave acesso nfe",
}

try:
    user32.SetProcessDPIAware()
except Exception:
    pass


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


user32.OpenClipboard.argtypes = [ctypes.c_void_p]
user32.OpenClipboard.restype = ctypes.c_int
user32.CloseClipboard.argtypes = []
user32.CloseClipboard.restype = ctypes.c_int
user32.EmptyClipboard.argtypes = []
user32.EmptyClipboard.restype = ctypes.c_int
user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
user32.SetClipboardData.restype = ctypes.c_void_p
user32.GetAsyncKeyState.argtypes = [ctypes.c_int]
user32.GetAsyncKeyState.restype = ctypes.c_short
user32.keybd_event.argtypes = [ctypes.c_ubyte, ctypes.c_ubyte, ctypes.c_uint, ctypes.c_ulong]
user32.keybd_event.restype = None
user32.GetCursorPos.argtypes = [ctypes.POINTER(POINT)]
user32.GetCursorPos.restype = ctypes.c_int
user32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
user32.SetCursorPos.restype = ctypes.c_int
user32.mouse_event.argtypes = [ctypes.c_uint, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint, ctypes.c_ulong]
user32.mouse_event.restype = None

kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
kernel32.GlobalAlloc.restype = ctypes.c_void_p
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
kernel32.GlobalUnlock.restype = ctypes.c_int
kernel32.GlobalFree.argtypes = [ctypes.c_void_p]
kernel32.GlobalFree.restype = ctypes.c_void_p


def key_down(vk: int) -> bool:
    return bool(user32.GetAsyncKeyState(vk) & 0x8000)


def sleep_abortable(seconds: float, stop_event: threading.Event) -> None:
    end = time.monotonic() + seconds
    while True:
        if stop_event.is_set() or key_down(VK_F8):
            raise KeyboardInterrupt
        remaining = end - time.monotonic()
        if remaining <= 0:
            return
        time.sleep(min(0.05, remaining))


def click_at(x: int, y: int) -> None:
    user32.SetCursorPos(x, y)
    user32.mouse_event(0x0002, 0, 0, 0, 0)
    user32.mouse_event(0x0004, 0, 0, 0, 0)


def hotkey_ctrl(vk: int) -> None:
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    user32.keybd_event(vk, 0, 0, 0)
    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)


def copiar_para_clipboard(texto: str) -> None:
    data = ctypes.create_unicode_buffer(texto)
    size_bytes = ctypes.sizeof(data)

    ultimo_erro = None
    for _ in range(10):
        if not user32.OpenClipboard(None):
            ultimo_erro = "Nao consegui abrir o clipboard."
            time.sleep(0.05)
            continue

        handle = None
        try:
            if not user32.EmptyClipboard():
                ultimo_erro = "Nao consegui limpar o clipboard."
                continue

            handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, size_bytes)
            if not handle:
                ultimo_erro = "Nao consegui alocar memoria para o clipboard."
                continue

            pointer = kernel32.GlobalLock(handle)
            if not pointer:
                ultimo_erro = "Nao consegui bloquear a memoria do clipboard."
                continue

            ctypes.memmove(pointer, data, size_bytes)
            kernel32.GlobalUnlock(handle)

            if not user32.SetClipboardData(CF_UNICODETEXT, handle):
                ultimo_erro = "Nao consegui definir o conteudo do clipboard."
                continue

            handle = None
            return
        finally:
            if handle:
                kernel32.GlobalFree(handle)
            user32.CloseClipboard()

    raise RuntimeError(ultimo_erro or "Falha ao copiar para o clipboard.")


def normalizar_texto(valor) -> str:
    if valor is None:
        return ""
    texto = str(valor).strip().lower()
    return re.sub(r"\s+", " ", texto)


def extrair_chave(valor) -> str | None:
    if valor is None:
        return None
    texto = str(valor).strip()
    if not texto:
        return None
    digitos = re.sub(r"\D", "", texto)
    if len(digitos) == 44:
        return digitos
    return None


def carregar_chaves_xlsx(caminho: Path) -> list[str]:
    def col_ref_para_indice(ref: str) -> int:
        idx = 0
        for ch in ref:
            if ch.isalpha():
                idx = idx * 26 + (ord(ch.upper()) - 64)
        return idx

    ns_main = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    ns_rel = {"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

    import zipfile

    with zipfile.ZipFile(caminho, "r") as zf:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            raiz_ss = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in raiz_ss.findall("a:si", ns_main):
                shared_strings.append("".join(t.text or "" for t in si.findall(".//a:t", ns_main)))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {}
        for rel in rels:
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            if rid and target:
                rel_map[rid] = target

        chaves: list[str] = []
        for sheet in workbook.findall("a:sheets/a:sheet", ns_main):
            rel_id = sheet.attrib.get(f"{{{ns_rel['r']}}}id")
            target = rel_map.get(rel_id)
            if not target:
                continue

            sheet_path = target if target.startswith("xl/") else f"xl/{target}"
            if sheet_path not in zf.namelist():
                continue

            raiz_sheet = ET.fromstring(zf.read(sheet_path))
            rows = []
            for row in raiz_sheet.findall(".//a:sheetData/a:row", ns_main):
                row_index = int(row.attrib.get("r", "0") or "0")
                cells = {}
                for cell in row.findall("a:c", ns_main):
                    ref = cell.attrib.get("r", "")
                    match = re.match(r"([A-Z]+)(\d+)", ref, re.I)
                    if not match:
                        continue
                    col_index = col_ref_para_indice(match.group(1))
                    tipo = cell.attrib.get("t", "")
                    if tipo == "s":
                        idx = cell.findtext("a:v", default="", namespaces=ns_main)
                        valor = shared_strings[int(idx)] if idx.isdigit() and int(idx) < len(shared_strings) else ""
                    elif tipo == "inlineStr":
                        valor = "".join(t.text or "" for t in cell.findall(".//a:t", ns_main))
                    else:
                        valor = cell.findtext("a:v", default="", namespaces=ns_main) or ""
                    cells[col_index] = valor
                rows.append((row_index, cells))

            if not rows:
                continue

            col_idx = 1
            start_row = 1
            for row_index, cells in rows[:20]:
                for col, valor in cells.items():
                    texto = normalizar_texto(valor)
                    if texto in NOMES_DE_COLUNA or any(alias in texto for alias in NOMES_DE_COLUNA):
                        col_idx = col
                        start_row = row_index + 1
                        break
                if start_row != 1:
                    break

            for row_index, cells in rows:
                if row_index < start_row:
                    continue
                chave = extrair_chave(cells.get(col_idx))
                if chave:
                    chaves.append(chave)

    return list(dict.fromkeys(chaves))


def carregar_chaves_csv(caminho: Path) -> list[str]:
    with caminho.open("r", encoding="utf-8-sig", newline="") as f:
        amostra = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(amostra, delimiters=";,\t|")
            reader = csv.reader(f, dialect)
        except csv.Error:
            reader = csv.reader(f, delimiter=";")

        linhas = list(reader)

    if not linhas:
        return []

    header = [normalizar_texto(col) for col in linhas[0]]
    col_idx = 0
    start = 0
    for i, name in enumerate(header):
        if name in NOMES_DE_COLUNA or any(alias in name for alias in NOMES_DE_COLUNA):
            col_idx = i
            start = 1
            break

    chaves: list[str] = []
    for linha in linhas[start:]:
        if col_idx < len(linha):
            chave = extrair_chave(linha[col_idx])
            if chave:
                chaves.append(chave)

    return list(dict.fromkeys(chaves))


def carregar_chaves(caminho: Path) -> list[str]:
    ext = caminho.suffix.lower()
    if ext not in EXTENSOES_SUPORTADAS:
        raise ValueError(f"Extensao nao suportada: {ext}")
    if ext in {".xlsx", ".xlsm"}:
        return carregar_chaves_xlsx(caminho)
    return carregar_chaves_csv(caminho)


def executar_fluxo(chaves: list[str], stop_event: threading.Event) -> None:
    total = len(chaves)
    for idx, chave in enumerate(chaves, start=1):
        if stop_event.is_set() or key_down(VK_F8):
            raise KeyboardInterrupt

        print(f"[{idx}/{total}] Processando chave...")
        click_at(*COORD_FIELD)
        sleep_abortable(PAUSA_DEPOIS_CAMPO, stop_event)
        copiar_para_clipboard(chave)
        hotkey_ctrl(VK_V)
        sleep_abortable(PAUSA_DEPOIS_COLAR, stop_event)

        click_at(*COORD_1)
        sleep_abortable(PAUSA_DEPOIS_COLAR, stop_event)
        click_at(*COORD_2)
        sleep_abortable(PAUSA_DEPOIS_COLAR, stop_event)
        click_at(*COORD_3)
        sleep_abortable(PAUSA_DEPOIS_COLAR, stop_event)
        click_at(*COORD_4)
        sleep_abortable(PAUSA_ENTRE_BLOCO, stop_event)


class App:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title("Fluxo fixo por coordenadas")
        self.root.geometry("560x240")
        self.root.resizable(False, False)

        self.status = StringVar(value="Selecione a planilha e execute o fluxo.")
        self.keys: list[str] = []
        self.stop_event = threading.Event()
        self.running = False

        self._build_ui()
        self._refresh()

    def _build_ui(self) -> None:
        import tkinter.ttk as ttk

        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Fluxo fixo por coordenadas", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(frame, textvariable=self.status, wraplength=520).pack(anchor="w", pady=(8, 10))

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(0, 10))

        ttk.Button(buttons, text="Selecionar planilha", command=self.select_sheet).pack(side="left")
        ttk.Button(buttons, text="Executar", command=self.start).pack(side="left", padx=8)
        ttk.Button(buttons, text="Cancelar", command=self.cancel).pack(side="left")

        self.summary = ttk.Label(frame, text="", justify="left")
        self.summary.pack(anchor="w", pady=(8, 0))

        ttk.Label(
            frame,
            text=(
                "Sequencia: clicar em 685,304 -> colar chave -> 1050,349 -> 272,435 -> 1096,417 -> 763,346. "
                "F8 cancela. Os delays ficam no topo do arquivo."
            ),
            wraplength=520,
        ).pack(anchor="w", pady=(12, 0))

    def _refresh(self) -> None:
        self.summary.configure(
            text=(
                f"Chaves carregadas: {len(self.keys)}\n"
                f"Coordenadas fixas: {COORD_FIELD} -> {COORD_1} -> {COORD_2} -> {COORD_3} -> {COORD_4}"
                f"\nDelays: campo={PAUSA_DEPOIS_CAMPO}s, colar={PAUSA_DEPOIS_COLAR}s, bloco={PAUSA_ENTRE_BLOCO}s"
            )
        )

    def set_status(self, text: str) -> None:
        self.status.set(text)

    def select_sheet(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecione a planilha com as chaves de acesso",
            filetypes=[
                ("Planilhas", "*.xlsx *.xlsm *.csv"),
                ("Excel", "*.xlsx *.xlsm"),
                ("CSV", "*.csv"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if not path:
            return

        try:
            self.keys = carregar_chaves(Path(path))
            if not self.keys:
                messagebox.showwarning("Sem chaves", "Nao encontrei chaves validas na planilha.")
            else:
                self.set_status(f"Planilha carregada com {len(self.keys)} chaves.")
        except Exception as exc:
            messagebox.showerror("Erro ao ler planilha", str(exc))
        self._refresh()

    def cancel(self) -> None:
        self.stop_event.set()
        self.set_status("Cancelando...")

    def start(self) -> None:
        if self.running:
            return
        if not self.keys:
            messagebox.showwarning("Sem planilha", "Selecione a planilha primeiro.")
            return

        self.stop_event.clear()
        self.running = True
        self.set_status("Iniciando em 3 segundos. Foque o navegador. F8 cancela.")
        self.root.withdraw()
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self) -> None:
        try:
            sleep_abortable(PAUSA_INICIAL, self.stop_event)
            executar_fluxo(self.keys, self.stop_event)
        except KeyboardInterrupt:
            pass
        except Exception as exc:
            self.root.after(0, lambda: messagebox.showerror("Erro na execucao", str(exc)))
        finally:
            self.running = False
            self.root.after(0, self._finish)

    def _finish(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(250, lambda: self.root.attributes("-topmost", False))
        self.set_status("Execucao finalizada ou cancelada.")


def main() -> None:
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
