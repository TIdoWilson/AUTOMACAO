import os
import re
import sys
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox


BASE_DIR = Path(r"W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\FAZEDOR DE AEF")
SCRIPT_1 = BASE_DIR / "1 - Baixador de Balancete de Verificacao.py"
SCRIPT_2 = BASE_DIR / "2 - Fromatador XLSX.py"
SCRIPT_3 = BASE_DIR / "3 - Organizar no Site.py"
ARQ_EMPRESAS = BASE_DIR / "empresas.txt"
ARQ_ENV = BASE_DIR / ".env"
PYTHON_ECD = Path(r"C:\Users\ECD\AppData\Local\Programs\Python\Python311\python.exe")


def ler_empresas_txt() -> list[str]:
    if not ARQ_EMPRESAS.exists():
        return []
    with ARQ_EMPRESAS.open("r", encoding="utf-8") as f:
        return [linha.strip().lstrip("\ufeff") for linha in f if linha.strip()]


def ler_mapa_codigo_nome_env() -> dict[str, str]:
    mapa: dict[str, str] = {}
    if not ARQ_ENV.exists():
        return mapa

    rgx = re.compile(r"^AEF_SITE_([A-Z0-9_]+)_CODIGO=(.+)$")
    with ARQ_ENV.open("r", encoding="utf-8") as f:
        for linha in f:
            m = rgx.match(linha.strip())
            if not m:
                continue
            perfil = m.group(1).replace("_", " ").strip()
            codigo = m.group(2).strip()
            if codigo:
                mapa[codigo] = perfil
    return mapa


def montar_opcoes_empresas() -> list[tuple[str, str]]:
    codigos = []
    for c in ler_empresas_txt():
        if c not in codigos:
            codigos.append(c)

    mapa_env = ler_mapa_codigo_nome_env()
    for c in mapa_env.keys():
        if c not in codigos:
            codigos.append(c)

    opcoes: list[tuple[str, str]] = []
    for codigo in codigos:
        nome = mapa_env.get(codigo, "SEM NOME")
        opcoes.append((codigo, nome))

    opcoes.sort(key=lambda x: (x[1], x[0]))
    return opcoes


class OrquestradorAEF:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Orquestrador AEF")
        self.root.geometry("900x620")

        self.executando = False
        self.opcoes = montar_opcoes_empresas()
        self.display_para_codigo: dict[str, str] = {}

        self._montar_tela()
        self._carregar_empresas_combo()

    def _montar_tela(self) -> None:
        frame_topo = ttk.Frame(self.root, padding=12)
        frame_topo.pack(fill="x")

        ttk.Label(frame_topo, text="Empresa (nome + codigo):").grid(row=0, column=0, sticky="w")
        self.var_empresa = tk.StringVar()
        self.combo_empresa = ttk.Combobox(
            frame_topo,
            textvariable=self.var_empresa,
            state="readonly",
            width=55,
        )
        self.combo_empresa.grid(row=1, column=0, sticky="w")

        ttk.Label(frame_topo, text="Filtro (nome ou codigo):").grid(row=0, column=1, padx=(12, 0), sticky="w")
        self.var_filtro = tk.StringVar()
        ent_filtro = ttk.Entry(frame_topo, textvariable=self.var_filtro, width=28)
        ent_filtro.grid(row=1, column=1, padx=(12, 0), sticky="w")
        ent_filtro.bind("<KeyRelease>", lambda _e: self._filtrar_empresas())

        frame_botoes = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        frame_botoes.pack(fill="x")

        self.btn_1e2 = ttk.Button(
            frame_botoes,
            text="Baixar e Formatar Balancete do ano",
            command=self.acao_rodar_1e2,
        )
        self.btn_1e2.pack(side="left")

        self.btn_3 = ttk.Button(
            frame_botoes,
            text="Enviar arquivos prontos para o site",
            command=self.acao_rodar_3,
        )
        self.btn_3.pack(side="left", padx=(8, 0))

        self.btn_recarregar = ttk.Button(frame_botoes, text="Recarregar Empresas", command=self.acao_recarregar)
        self.btn_recarregar.pack(side="left", padx=(8, 0))

        frame_alertas = ttk.LabelFrame(self.root, text="Avisos Operacionais", padding=12)
        frame_alertas.pack(fill="x", padx=12, pady=(0, 8))
        ttk.Label(
            frame_alertas,
            text=(
                "1) Script 1 controla mouse e teclado. Nao use o computador durante a execucao.\n"
                "2) Script 3 abre navegador (Chrome/Chromium) para login e lancamento no site."
            ),
            justify="left",
        ).pack(anchor="w")

        frame_log = ttk.LabelFrame(self.root, text="Log", padding=8)
        frame_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.txt_log = tk.Text(frame_log, wrap="word", height=20)
        self.txt_log.pack(fill="both", expand=True)
        self._log("UI iniciada.")

    def _filtrar_empresas(self) -> None:
        termo = self.var_filtro.get().strip().lower()
        if not termo:
            filtradas = self.opcoes
        else:
            filtradas = [
                (codigo, nome)
                for codigo, nome in self.opcoes
                if termo in codigo.lower() or termo in nome.lower()
            ]
        self._carregar_empresas_combo(filtradas)

    def _carregar_empresas_combo(self, opcoes: list[tuple[str, str]] | None = None) -> None:
        if opcoes is None:
            opcoes = self.opcoes

        self.display_para_codigo.clear()
        displays: list[str] = []
        for codigo, nome in opcoes:
            display = f"{nome} ({codigo})"
            displays.append(display)
            self.display_para_codigo[display] = codigo

        self.combo_empresa["values"] = displays
        if displays:
            self.combo_empresa.current(0)
        else:
            self.var_empresa.set("")

    def _log(self, texto: str) -> None:
        self.txt_log.insert("end", texto + "\n")
        self.txt_log.see("end")

    def _set_executando(self, executando: bool) -> None:
        self.executando = executando
        estado = "disabled" if executando else "normal"
        self.btn_1e2.configure(state=estado)
        self.btn_3.configure(state=estado)
        self.btn_recarregar.configure(state=estado)
        self.combo_empresa.configure(state="disabled" if executando else "readonly")

    def _python_exec(self) -> str:
        if PYTHON_ECD.exists():
            return str(PYTHON_ECD)
        return sys.executable

    def _empresa_selecionada(self) -> str | None:
        display = self.var_empresa.get().strip()
        codigo = self.display_para_codigo.get(display)
        if not codigo:
            messagebox.showerror("Erro", "Selecione uma empresa valida.")
            return None
        return codigo

    def _rodar_comando(self, cmd: list[str], titulo: str) -> int:
        self._log(f"[{titulo}] Comando: {' '.join(cmd)}")
        proc = subprocess.Popen(
            cmd,
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        assert proc.stdout is not None
        for linha in proc.stdout:
            self._log(linha.rstrip("\r\n"))
        return proc.wait()

    def _rodar_em_thread(self, alvo) -> None:
        if self.executando:
            messagebox.showwarning("Em execucao", "Ja existe uma execucao em andamento.")
            return

        self._set_executando(True)

        def runner():
            try:
                alvo()
            except Exception as exc:
                self.root.after(0, lambda: self._log(f"ERRO: {exc}"))
                self.root.after(0, lambda: messagebox.showerror("Erro", str(exc)))
            finally:
                self.root.after(0, lambda: self._set_executando(False))

        threading.Thread(target=runner, daemon=True).start()

    def acao_recarregar(self) -> None:
        self.opcoes = montar_opcoes_empresas()
        self._carregar_empresas_combo()
        self._log("Empresas recarregadas.")

    def _sobrescrever_empresas_txt_temporario(self, codigo: str):
        if not ARQ_EMPRESAS.exists():
            raise RuntimeError(f"Arquivo nao encontrado: {ARQ_EMPRESAS}")
        original = ARQ_EMPRESAS.read_text(encoding="utf-8")
        ARQ_EMPRESAS.write_text(codigo + "\n", encoding="utf-8")
        return original

    def acao_rodar_1e2(self) -> None:
        codigo = self._empresa_selecionada()
        if not codigo:
            return

        ok = messagebox.askyesno(
            "Confirmacao",
            (
                "O Script 1 controla mouse/teclado e pode interromper seu uso do computador.\n\n"
                f"Executar 'Baixar e Formatar Balancete do ano' para a empresa {codigo}?"
            ),
        )
        if not ok:
            return

        def fluxo():
            py = self._python_exec()
            self._log(f"Iniciando Script 1 + 2 para empresa {codigo}.")

            backup_empresas = self._sobrescrever_empresas_txt_temporario(codigo)
            try:
                rc1 = self._rodar_comando([py, str(SCRIPT_1)], "SCRIPT 1")
                self._log(f"[SCRIPT 1] Retorno: {rc1}")
                if rc1 != 0:
                    self._log("Script 1 falhou. Script 2 nao sera executado.")
                    return

                rc2 = self._rodar_comando([py, str(SCRIPT_2), "--empresa", codigo], "SCRIPT 2")
                self._log(f"[SCRIPT 2] Retorno: {rc2}")
                if rc2 == 0:
                    self._log("Fluxo 1 -> 2 finalizado com sucesso.")
                else:
                    self._log("Script 2 finalizou com erro.")
            finally:
                ARQ_EMPRESAS.write_text(backup_empresas, encoding="utf-8")
                self._log("empresas.txt restaurado.")

        self._rodar_em_thread(fluxo)

    def acao_rodar_3(self) -> None:
        codigo = self._empresa_selecionada()
        if not codigo:
            return

        ok = messagebox.askyesno(
            "Confirmacao",
            (
                "O Script 3 abre navegador (Chrome/Chromium) e executa automacao no site.\n\n"
                f"Executar 'Enviar arquivos prontos para o site' para a empresa {codigo}?"
            ),
        )
        if not ok:
            return

        def fluxo():
            py = self._python_exec()
            self._log(f"Iniciando Script 3 para empresa {codigo}.")
            rc3 = self._rodar_comando([py, str(SCRIPT_3), "--empresa", codigo], "SCRIPT 3")
            self._log(f"[SCRIPT 3] Retorno: {rc3}")
            if rc3 == 0:
                self._log("Script 3 finalizado com sucesso.")
            else:
                self._log("Script 3 finalizou com erro.")

        self._rodar_em_thread(fluxo)


def main() -> int:
    if not BASE_DIR.exists():
        print(f"ERRO: pasta base nao encontrada: {BASE_DIR}")
        return 1

    root = tk.Tk()
    app = OrquestradorAEF(root)
    if not PYTHON_ECD.exists():
        app._log("AVISO: Python ECD nao encontrado. Usando python atual (sys.executable).")
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
