# -*- coding: utf-8 -*-
"""
Conversor em lote: PDF -> PDF/A (Windows)

- UI simples (Tkinter) para selecionar uma pasta.
- Varre subpastas e converte todos os PDFs para PDF/A, sobrescrevendo o original
  somente se a conversao gerar o arquivo de saida com sucesso.
- Usa Ghostscript (gswin64c.exe).

Observacao importante:
Converter para PDF/A nao mantem o arquivo "idêntico por bytes". O objetivo e
preservar o conteudo visual e gerar conformidade PDF/A; dependendo de fontes,
transparencias e outros recursos, o arquivo pode mudar internamente.
"""

from __future__ import annotations

import os
import sys
import time
import queue
import shutil
import threading
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox


# ========= CONFIG =========
WINDOW_TITLE = "Conversor PDF -> PDF/A (Ghostscript)"

# Pastas ignoradas durante o walk (para reduzir risco de varrer lixo/artefatos).
DEFAULT_IGNORE_DIRS = {
    "__pycache__",
    "bin",
    "obj",
    "publish",
    "venv",
    "venvs_servidor",
}

# Onde salvar logs (ao lado deste script).
LOG_DIR = Path(__file__).resolve().parent

# ========= FIM CONFIG =========


@dataclass(frozen=True)
class ConvertResult:
    ok: bool
    input_path: Path
    message: str


def is_probably_pdfa(pdf_path: Path, max_bytes: int = 8 * 1024 * 1024) -> bool:
    """
    Heuristica (nao e validacao formal):
    - procura marcadores comuns em PDF/A no conteudo binario.
    - se encontrar os principais, considera "ja e PDF/A" e podemos pular.

    Preferimos falso-negativo (reconverter) do que falso-positivo (pular sem ser PDF/A).
    """
    markers = [
        b"pdfaid:part",
        b"pdfaid:conformance",
        b"/OutputIntents",
        b"GTS_PDFA",
    ]
    found = {m: False for m in markers}

    try:
        total = pdf_path.stat().st_size
        to_read = min(total, max_bytes)
        chunk_size = 1024 * 1024
        carry = b""
        read_so_far = 0

        with pdf_path.open("rb") as f:
            while read_so_far < to_read:
                n = min(chunk_size, to_read - read_so_far)
                chunk = f.read(n)
                if not chunk:
                    break
                read_so_far += len(chunk)

                hay = carry + chunk
                for m in markers:
                    if not found[m] and m in hay:
                        found[m] = True

                # Mantem uma "janela" pequena para pegarmos matches que cruzam o limite do chunk.
                carry = hay[-4096:]

                if all(found.values()):
                    return True
    except Exception:
        return False

    # Regra conservadora: so considera PDF/A se achou os 3 principais.
    return found[b"pdfaid:part"] and found[b"pdfaid:conformance"] and found[b"/OutputIntents"]


def gs_extract_text_len(gs_exe: str, pdf_path: Path, timeout_s: int = 45) -> tuple[bool, int, str]:
    """
    Extrai texto via Ghostscript (txtwrite). Serve para comparar "input tinha texto" vs "output virou imagem".
    Retorna: (ok, tamanho_texto, msg).
    """
    try:
        with tempfile.TemporaryDirectory(prefix="pdfa_txt_") as td:
            out_txt = Path(td) / "out.txt"
            args = [
                gs_exe,
                "-dNOSAFER",
                "-dBATCH",
                "-dNOPAUSE",
                "-sDEVICE=txtwrite",
                "-o",
                str(out_txt),
                "-f",
                str(pdf_path),
            ]
            cp = subprocess.run(
                args,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout_s,
            )
            if cp.returncode != 0:
                err = (cp.stderr or "").strip()
                out = (cp.stdout or "").strip()
                msg = f"txtwrite retornou {cp.returncode}"
                if err:
                    msg += f" | stderr: {err[:2000]}"
                if out:
                    msg += f" | stdout: {out[:2000]}"
                return False, 0, msg

            if not out_txt.exists():
                return False, 0, "txtwrite nao gerou arquivo de saida"

            # Heuristica de "texto real": conta caracteres alfanumericos (evita false-positive com whitespace).
            data = out_txt.read_text(encoding="utf-8", errors="replace")
            alnum = sum(1 for ch in data if ch.isalnum())
            return True, alnum, "OK"
    except subprocess.TimeoutExpired:
        return False, 0, f"Timeout ({timeout_s}s) no txtwrite"
    except FileNotFoundError:
        return False, 0, "Ghostscript nao encontrado (gswin64c.exe)"
    except Exception as e:
        return False, 0, f"Erro no txtwrite: {e}"


def _now_stamp() -> str:
    return time.strftime("%Y%m%d_%H%M%S")


def log_path_for_run() -> Path:
    return LOG_DIR / f"pdfa_converter_log_{_now_stamp()}.txt"


def append_log(log_path: Path, line: str) -> None:
    # Nao falhar a conversao por causa de log.
    try:
        with log_path.open("a", encoding="utf-8", errors="replace") as f:
            f.write(line.rstrip("\n") + "\n")
    except Exception:
        pass


def which_gswin64c() -> str | None:
    return shutil.which("gswin64c") or shutil.which("gswin64c.exe")


def _candidate_paths_near_gs(gs_exe: Path, rel_parts: list[str]) -> list[Path]:
    p = gs_exe.resolve()
    out: list[Path] = []
    for up in (0, 1, 2, 3):
        base = p.parent
        for _ in range(up):
            base = base.parent
        out.append(base.joinpath(*rel_parts))
    return out


def find_pdfa_def_ps(gs_exe: str) -> Path | None:
    """
    Tenta localizar PDFA_def.ps na instalacao do Ghostscript.
    """
    p = Path(gs_exe)
    candidates = []
    candidates += _candidate_paths_near_gs(p, ["lib", "PDFA_def.ps"])
    candidates += _candidate_paths_near_gs(p, ["Resource", "Init", "PDFA_def.ps"])
    for c in candidates:
        if c.exists():
            return c

    # Fallback: parse de "gswin64c -h" para tentar achar Resource\\Init.
    try:
        cp = subprocess.run(
            [gs_exe, "-h"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
        )
        text_out = (cp.stdout or "") + "\n" + (cp.stderr or "")
        for token in text_out.splitlines():
            # linhas do help variam; procuramos qualquer caminho que contenha Resource\\Init
            if "Resource\\Init" in token or "Resource/Init" in token:
                # extrai possiveis paths separados por ';'
                for part in token.replace(";", "\n").splitlines():
                    part = part.strip().strip('"')
                    if not part:
                        continue
                    if "Resource\\Init" in part or "Resource/Init" in part:
                        base = Path(part)
                        cand = base / "PDFA_def.ps"
                        if cand.exists():
                            return cand
    except Exception:
        return None
    return None


def find_srgb_icc(gs_exe: str) -> Path | None:
    """
    Opcional: tenta localizar sRGB.icc para ajudar na conformidade PDF/A.
    Se nao achar, o Ghostscript pode ainda conseguir pelo search path.
    """
    p = Path(gs_exe)
    candidates = []
    candidates += _candidate_paths_near_gs(p, ["iccprofiles", "sRGB.icc"])
    candidates += _candidate_paths_near_gs(p, ["lib", "sRGB.icc"])
    candidates += _candidate_paths_near_gs(p, ["Resource", "ICCProfiles", "sRGB.icc"])
    for c in candidates:
        if c.exists():
            return c
    return None


def list_pdfs(root_dir: Path, ignore_dirs: set[str]) -> list[Path]:
    pdfs: list[Path] = []
    root_dir = root_dir.resolve()
    for dirpath, dirnames, filenames in os.walk(root_dir):
        # filtra dirs in-place
        dirnames[:] = [d for d in dirnames if d not in ignore_dirs]
        for name in filenames:
            if name.lower().endswith(".pdf"):
                pdfs.append(Path(dirpath) / name)
    return pdfs


def run_gs_convert_to_pdfa(
    gs_exe: str,
    input_pdf: Path,
    output_tmp: Path,
    pdfa_def_ps: Path | None,
    srgb_icc: Path | None,
    pdfa_level: int = 2,
    timeout_s: int = 300,
) -> tuple[bool, str]:
    """
    Converte um arquivo para PDF/A via Ghostscript.
    Retorna (ok, mensagem).
    """
    # Observacao:
    # - PDF/A-1 e mais restritivo e pode "achatar" recursos (ex.: transparencia) e, em alguns casos,
    #   acabar gerando saida menos editavel/selecionavel.
    # - PDF/A-2 tende a preservar melhor (recomendado).
    if pdfa_level not in (1, 2):
        pdfa_level = 2

    if not srgb_icc or not Path(srgb_icc).exists():
        return False, "Perfil ICC sRGB nao encontrado (necessario para PDF/A)."

    args: list[str] = [
        gs_exe,
        # Ferramenta local (assistida): o modo SAFER/permit pode quebrar com caminhos com espacos/UNC.
        # Para evitar isso, usamos NOSAFER.
        "-dNOSAFER",
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        "-sDEVICE=pdfwrite",
        f"-dPDFA={pdfa_level}",
        "-dPDFACompatibilityPolicy=1",
        "-sProcessColorModel=DeviceRGB",
        "-sColorConversionStrategy=RGB",
        f"-sOutputICCProfile={str(srgb_icc)}",
        "-o",
        str(output_tmp),
    ]

    # O pdfwrite consegue gerar PDF/A usando -sOutputICCProfile.
    # Mantemos o PDFA_def.ps apenas para referencia/log (nao necessario aqui).
    args += ["-f", str(input_pdf)]

    try:
        cp = subprocess.run(
            args,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_s,
        )
    except subprocess.TimeoutExpired:
        return False, f"Timeout ({timeout_s}s) no Ghostscript"
    except FileNotFoundError:
        return False, "Ghostscript nao encontrado (gswin64c.exe)"
    except Exception as e:
        return False, f"Erro executando Ghostscript: {e}"

    out = (cp.stdout or "").strip()
    err = (cp.stderr or "").strip()
    if cp.returncode != 0:
        msg = f"Ghostscript retornou {cp.returncode}"
        # Em alguns casos o GS loga em stdout; em outros, em stderr.
        if err:
            msg += f" | stderr: {err[:8000]}"
        if out:
            msg += f" | stdout: {out[:8000]}"
        return False, msg

    if not output_tmp.exists() or output_tmp.stat().st_size == 0:
        return False, "Arquivo de saida nao foi gerado (ou ficou vazio)"

    return True, "OK"


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(WINDOW_TITLE)
        self.geometry("880x520")

        self._stop_event = threading.Event()
        self._worker: threading.Thread | None = None
        self._q: queue.Queue[str] = queue.Queue()
        self._log_path = log_path_for_run()

        self.var_root = tk.StringVar(value="")
        self.var_gs = tk.StringVar(value=which_gswin64c() or "")
        self.var_ignore = tk.StringVar(value=";".join(sorted(DEFAULT_IGNORE_DIRS)))
        self.var_skip_pdfa = tk.BooleanVar(value=True)
        self.var_pdfa_level = tk.StringVar(value="2")
        self.var_check_text = tk.BooleanVar(value=True)

        self.var_total = tk.StringVar(value="0")
        self.var_ok = tk.StringVar(value="0")
        self.var_fail = tk.StringVar(value="0")
        self.var_skip = tk.StringVar(value="0")

        self._build_ui()
        self.after(100, self._drain_queue)

    def _build_ui(self) -> None:
        pad = 8

        frm = tk.Frame(self)
        frm.pack(fill="both", expand=True, padx=pad, pady=pad)

        # Pasta
        row1 = tk.Frame(frm)
        row1.pack(fill="x")
        tk.Label(row1, text="Pasta raiz:").pack(side="left")
        tk.Entry(row1, textvariable=self.var_root).pack(side="left", fill="x", expand=True, padx=(pad, pad))
        tk.Button(row1, text="Selecionar...", command=self._pick_root).pack(side="left")

        # Ghostscript
        row2 = tk.Frame(frm)
        row2.pack(fill="x", pady=(pad, 0))
        tk.Label(row2, text="Ghostscript (gswin64c.exe):").pack(side="left")
        tk.Entry(row2, textvariable=self.var_gs).pack(side="left", fill="x", expand=True, padx=(pad, pad))
        tk.Button(row2, text="Localizar...", command=self._pick_gs).pack(side="left")

        # Ignorar pastas
        row3 = tk.Frame(frm)
        row3.pack(fill="x", pady=(pad, 0))
        tk.Label(row3, text="Ignorar pastas (separado por ';'):").pack(side="left")
        tk.Entry(row3, textvariable=self.var_ignore).pack(side="left", fill="x", expand=True, padx=(pad, 0))

        # Acoes
        row4 = tk.Frame(frm)
        row4.pack(fill="x", pady=(pad, 0))
        self.btn_start = tk.Button(row4, text="Iniciar (sobrescreve originais)", command=self._start)
        self.btn_start.pack(side="left")
        self.btn_cancel = tk.Button(row4, text="Cancelar", command=self._cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=(pad, 0))
        tk.Checkbutton(row4, text="Pular PDFs ja PDF/A", variable=self.var_skip_pdfa).pack(side="left", padx=(pad, 0))
        tk.Checkbutton(
            row4,
            text="Nao sobrescrever se perder texto",
            variable=self.var_check_text,
        ).pack(side="left", padx=(pad, 0))
        tk.Label(row4, text="PDF/A:").pack(side="left", padx=(pad, 0))
        tk.Radiobutton(row4, text="2 (recomendado)", variable=self.var_pdfa_level, value="2").pack(side="left")
        tk.Radiobutton(row4, text="1", variable=self.var_pdfa_level, value="1").pack(side="left")
        tk.Button(row4, text="Abrir log", command=self._open_log).pack(side="right")

        # Contadores
        row5 = tk.Frame(frm)
        row5.pack(fill="x", pady=(pad, 0))
        tk.Label(row5, text="Total:").pack(side="left")
        tk.Label(row5, textvariable=self.var_total, width=8, anchor="w").pack(side="left")
        tk.Label(row5, text="OK:").pack(side="left", padx=(pad, 0))
        tk.Label(row5, textvariable=self.var_ok, width=8, anchor="w").pack(side="left")
        tk.Label(row5, text="Falhas:").pack(side="left", padx=(pad, 0))
        tk.Label(row5, textvariable=self.var_fail, width=8, anchor="w").pack(side="left")
        tk.Label(row5, text="Pulos:").pack(side="left", padx=(pad, 0))
        tk.Label(row5, textvariable=self.var_skip, width=8, anchor="w").pack(side="left")

        # Log visual
        row6 = tk.Frame(frm)
        row6.pack(fill="both", expand=True, pady=(pad, 0))
        tk.Label(row6, text="Status:").pack(anchor="w")
        self.txt = tk.Text(row6, height=18, wrap="none")
        self.txt.pack(fill="both", expand=True)
        self.txt.configure(state="disabled")

        self._ui_log(f"Log: {self._log_path}")

    def _ui_log(self, msg: str) -> None:
        self.txt.configure(state="normal")
        self.txt.insert("end", msg.rstrip("\n") + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _pick_root(self) -> None:
        p = filedialog.askdirectory(title="Selecione a pasta raiz")
        if p:
            self.var_root.set(p)

    def _pick_gs(self) -> None:
        p = filedialog.askopenfilename(
            title="Selecione o gswin64c.exe",
            filetypes=[("Ghostscript", "gswin64c.exe"), ("Executavel", "*.exe"), ("Todos", "*.*")],
        )
        if p:
            self.var_gs.set(p)

    def _open_log(self) -> None:
        try:
            if not self._log_path.exists():
                append_log(self._log_path, "# log criado (vazio)")
            os.startfile(str(self._log_path))  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("Erro", f"Nao foi possivel abrir o log: {e}")

    def _set_running(self, running: bool) -> None:
        self.btn_start.configure(state=("disabled" if running else "normal"))
        self.btn_cancel.configure(state=("normal" if running else "disabled"))

    def _cancel(self) -> None:
        self._stop_event.set()
        self._ui_log("Cancelamento solicitado. Aguardando terminar o arquivo atual...")
        append_log(self._log_path, "CANCELAMENTO SOLICITADO")

    def _start(self) -> None:
        if self._worker and self._worker.is_alive():
            return

        root = self.var_root.get().strip()
        gs = self.var_gs.get().strip()
        if not root:
            messagebox.showwarning("Aviso", "Selecione a pasta raiz.")
            return
        if not gs:
            messagebox.showwarning("Aviso", "Informe o caminho do gswin64c.exe (Ghostscript).")
            return

        root_dir = Path(root)
        if not root_dir.exists():
            messagebox.showerror("Erro", f"Pasta nao existe: {root_dir}")
            return

        append_log(self._log_path, f"ROOT={root_dir}")
        append_log(self._log_path, f"GS={gs}")
        append_log(self._log_path, f"IGNORE={self.var_ignore.get().strip()}")

        self._stop_event.clear()
        self.var_total.set("0")
        self.var_ok.set("0")
        self.var_fail.set("0")
        self.var_skip.set("0")
        self._set_running(True)

        self._worker = threading.Thread(target=self._worker_run, args=(root_dir, gs), daemon=True)
        self._worker.start()

    def _worker_run(self, root_dir: Path, gs_exe: str) -> None:
        ignore = {p.strip() for p in self.var_ignore.get().split(";") if p.strip()}
        skip_pdfa = bool(self.var_skip_pdfa.get())
        check_text = bool(self.var_check_text.get())
        try:
            pdfa_level = int(self.var_pdfa_level.get() or "2")
        except Exception:
            pdfa_level = 2

        pdfa_def = find_pdfa_def_ps(gs_exe)
        srgb_icc = find_srgb_icc(gs_exe)

        if not pdfa_def:
            self._q.put("ERRO: nao encontrei PDFA_def.ps perto do Ghostscript. Se a conversao falhar, instale/ajuste o Ghostscript e aponte o gswin64c.exe correto.")
            append_log(self._log_path, "PDFA_def.ps NAO ENCONTRADO")
        else:
            self._q.put(f"PDFA_def.ps: {pdfa_def}")
            append_log(self._log_path, f"PDFA_def.ps={pdfa_def}")

        if srgb_icc:
            self._q.put(f"sRGB.icc: {srgb_icc}")
            append_log(self._log_path, f"sRGB.icc={srgb_icc}")
        else:
            append_log(self._log_path, "sRGB.icc NAO ENCONTRADO (seguindo assim mesmo)")

        pdfs = list_pdfs(root_dir, ignore)
        self._q.put(f"Encontrados {len(pdfs)} PDFs.")
        append_log(self._log_path, f"TOTAL_PDFS={len(pdfs)}")

        ok = 0
        fail = 0
        skipped = 0

        for idx, pdf in enumerate(pdfs, start=1):
            if self._stop_event.is_set():
                self._q.put("Cancelado pelo usuario.")
                append_log(self._log_path, "CANCELADO")
                break

            if skip_pdfa and is_probably_pdfa(pdf):
                skipped += 1
                self._q.put(f"[{idx}/{len(pdfs)}] SKIP (ja parece PDF/A): {pdf}")
                append_log(self._log_path, f"SKIP_PDFA: {pdf}")
                self._q.put(f"__COUNTS__ total={len(pdfs)} ok={ok} fail={fail} skip={skipped}")
                continue

            # Se o input tinha texto extraivel, nao queremos que o output "vire imagem" (sem texto).
            input_text_alnum = 0
            input_text_ok = True
            if check_text:
                ok_txt, alnum, msg_txt = gs_extract_text_len(gs_exe, pdf)
                if ok_txt:
                    input_text_alnum = alnum
                else:
                    input_text_ok = False
                    self._q.put(f"[{idx}/{len(pdfs)}] Aviso: nao consegui medir texto do input ({msg_txt}). Seguindo conversao...")
                    append_log(self._log_path, f"WARN_INPUT_TEXT: {pdf} | {msg_txt}")

            self._q.put(f"[{idx}/{len(pdfs)}] Convertendo: {pdf}")
            append_log(self._log_path, f"IN: {pdf}")

            tmp = pdf.with_name(pdf.name + ".tmp_pdfa")
            if tmp.exists():
                try:
                    tmp.unlink()
                except Exception:
                    pass

            ok_one, msg = run_gs_convert_to_pdfa(
                gs_exe=gs_exe,
                input_pdf=pdf,
                output_tmp=tmp,
                pdfa_def_ps=pdfa_def,
                srgb_icc=srgb_icc,
                pdfa_level=pdfa_level,
            )

            if ok_one:
                if check_text and input_text_ok and input_text_alnum >= 20:
                    ok_txt_out, out_alnum, msg_out = gs_extract_text_len(gs_exe, tmp)
                    if (not ok_txt_out) or out_alnum < 20:
                        ok_one = False
                        msg = f"Perda de texto detectada (input_alnum={input_text_alnum}, out_alnum={out_alnum}). Nao sobrescreveu."
                        if not ok_txt_out:
                            msg += f" | txt_out_err={msg_out}"
                        append_log(self._log_path, f"FALHA_TEXTO: {pdf} | {msg}")
                        try:
                            if tmp.exists():
                                tmp.unlink()
                        except Exception:
                            pass

            if ok_one:
                try:
                    os.replace(str(tmp), str(pdf))
                except Exception as e:
                    ok_one = False
                    msg = f"Falha ao sobrescrever o original: {e}"
                    try:
                        if tmp.exists():
                            tmp.unlink()
                    except Exception:
                        pass

            if ok_one:
                ok += 1
                self._q.put(f"  OK: {pdf}")
                append_log(self._log_path, f"OK: {pdf}")
            else:
                fail += 1
                self._q.put(f"  FALHA: {pdf} | {msg}")
                append_log(self._log_path, f"FALHA: {pdf} | {msg}")
                try:
                    if tmp.exists():
                        tmp.unlink()
                except Exception:
                    pass

            self._q.put(f"__COUNTS__ total={len(pdfs)} ok={ok} fail={fail} skip={skipped}")

        self._q.put("__DONE__")

    def _drain_queue(self) -> None:
        try:
            while True:
                msg = self._q.get_nowait()
                if msg.startswith("__COUNTS__"):
                    parts = dict(p.split("=", 1) for p in msg.replace("__COUNTS__", "").strip().split())
                    self.var_total.set(parts.get("total", self.var_total.get()))
                    self.var_ok.set(parts.get("ok", self.var_ok.get()))
                    self.var_fail.set(parts.get("fail", self.var_fail.get()))
                    self.var_skip.set(parts.get("skip", self.var_skip.get()))
                elif msg == "__DONE__":
                    self._set_running(False)
                    self._ui_log("Finalizado.")
                    append_log(self._log_path, "FINALIZADO")
                else:
                    self._ui_log(msg)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._drain_queue)


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
