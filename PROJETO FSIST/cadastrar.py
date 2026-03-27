import os
import re
import sys
import sqlite3
import shutil
import json
import subprocess
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ==========================
# cryptography (para ler dados do certificado)
# ==========================
try:
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography.hazmat.backends import default_backend
    from cryptography.x509.oid import NameOID, ObjectIdentifier
    CRYPTO_OK = True
except ImportError:
    CRYPTO_OK = False


def ler_dados_certificado_completo(caminho: str, senha: str):
    """
    Lê PFX/P12 e retorna (nome, cnpj, vencimento_dd/mm/aaaa).
    Se não conseguir, retorna strings vazias.
    """
    if not CRYPTO_OK:
        return "", "", ""

    try:
        with open(caminho, "rb") as f:
            data = f.read()

        pwd = senha.encode() if senha else None
        key, cert, others = pkcs12.load_key_and_certificates(data, pwd)
        if cert is None:
            return "", "", ""

        subject = cert.subject
        nome = ""
        cnpj = ""

        # Nome: COMMON_NAME
        try:
            cn_attrs = subject.get_attributes_for_oid(NameOID.COMMON_NAME)
            if cn_attrs:
                nome = cn_attrs[0].value
        except Exception:
            pass

        # Tentativas de CNPJ:
        # 1) OID ICP-Brasil 2.16.76.1.3.3
        try:
            oid_cnpj = ObjectIdentifier("2.16.76.1.3.3")
            cnpj_attrs = subject.get_attributes_for_oid(oid_cnpj)
            if cnpj_attrs and not cnpj:
                val = cnpj_attrs[0].value
                cnpj = "".join(filter(str.isdigit, val))
        except Exception:
            pass

        # 2) SERIAL_NUMBER
        if not cnpj:
            try:
                serial_attrs = subject.get_attributes_for_oid(NameOID.SERIAL_NUMBER)
                if serial_attrs:
                    val = serial_attrs[0].value
                    cnpj = "".join(filter(str.isdigit, val))
            except Exception:
                pass

        # 3) Regex procurando 14 dígitos no subject (caso nada acima funcione)
        if not cnpj:
            try:
                subject_str = subject.rfc4514_string()
                m = re.search(r"\d{14}", subject_str)
                if m:
                    cnpj = m.group(0)
            except Exception:
                pass

        venc = cert.not_valid_after.strftime("%d/%m/%Y")
        return nome, cnpj, venc
    except Exception:
        return "", "", ""


# ==========================
# Utilitários gerais
# ==========================

def get_script_dir() -> str:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ==========================
# Banco de dados
# ==========================

def init_db(conn: sqlite3.Connection):
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            cnpj TEXT NOT NULL UNIQUE,
            caminho_certificado TEXT NOT NULL,
            senha_certificado TEXT,
            nsu_atual INTEGER NOT NULL DEFAULT 0,
            ativo INTEGER NOT NULL DEFAULT 1,
            criado_em TEXT,
            atualizado_em TEXT
        );
        """
    )
    conn.commit()

    # Garante coluna vencimento_certificado
    try:
        cur.execute("ALTER TABLE empresas ADD COLUMN vencimento_certificado TEXT;")
        conn.commit()
    except sqlite3.OperationalError:
        pass


def carregar_empresas(conn: sqlite3.Connection, filtro: str | None = None):
    cur = conn.cursor()
    if filtro:
        like = f"%{filtro}%"
        cur.execute(
            """
            SELECT id, nome, cnpj, caminho_certificado, senha_certificado,
                   nsu_atual, ativo, criado_em, atualizado_em, vencimento_certificado
            FROM empresas
            WHERE nome LIKE ? OR cnpj LIKE ?
            ORDER BY id;
            """,
            (like, like),
        )
    else:
        cur.execute(
            """
            SELECT id, nome, cnpj, caminho_certificado, senha_certificado,
                   nsu_atual, ativo, criado_em, atualizado_em, vencimento_certificado
            FROM empresas
            ORDER BY id;
            """
        )
    return cur.fetchall()


def _backup_certificado(caminho_original: str, cnpj: str) -> str:
    """
    Copia o certificado para a pasta CERTIFICADOS ao lado da database
    e retorna o novo caminho.
    """
    base_dir = get_script_dir()
    cert_dir = os.path.join(base_dir, "CERTIFICADOS")
    os.makedirs(cert_dir, exist_ok=True)

    cnpj_num = "".join(filter(str.isdigit, cnpj))
    ext = os.path.splitext(caminho_original)[1] or ".pfx"
    destino = os.path.join(cert_dir, f"{cnpj_num}{ext}")

    shutil.copy2(caminho_original, destino)
    return destino


def inserir_empresa(conn: sqlite3.Connection, nome, cnpj, caminho, senha, nsu_inicial, vencimento):
    cur = conn.cursor()
    try:
        nsu_int = int("".join(filter(str.isdigit, str(nsu_inicial)))) if nsu_inicial else 0
    except ValueError:
        nsu_int = 0

    cnpj_num = "".join(filter(str.isdigit, cnpj))

    backup_path = _backup_certificado(caminho, cnpj_num)

    agora = datetime.now().isoformat(timespec="seconds")
    cur.execute(
        """
        INSERT INTO empresas (nome, cnpj, caminho_certificado, senha_certificado,
                              nsu_atual, ativo, criado_em, atualizado_em,
                              vencimento_certificado)
        VALUES (?, ?, ?, ?, ?, 1, ?, ?, ?);
        """,
        (nome.strip(), cnpj_num, backup_path, senha, nsu_int, agora, agora, vencimento.strip()),
    )
    conn.commit()


def atualizar_empresa(conn: sqlite3.Connection, empresa_id: int, nome, cnpj, caminho, senha, nsu_inicial, vencimento, ativo: int):
    cur = conn.cursor()
    try:
        nsu_int = int("".join(filter(str.isdigit, str(nsu_inicial)))) if nsu_inicial else 0
    except ValueError:
        nsu_int = 0

    cnpj_num = "".join(filter(str.isdigit, cnpj))

    cur.execute("SELECT caminho_certificado FROM empresas WHERE id = ?;", (empresa_id,))
    row = cur.fetchone()
    caminho_atual = row[0] if row else ""

    novo_caminho = caminho_atual
    if caminho and os.path.isfile(caminho) and os.path.abspath(caminho) != os.path.abspath(caminho_atual):
        novo_caminho = _backup_certificado(caminho, cnpj_num)

    agora = datetime.now().isoformat(timespec="seconds")
    cur.execute(
        """
        UPDATE empresas
           SET nome = ?,
               cnpj = ?,
               caminho_certificado = ?,
               senha_certificado = ?,
               nsu_atual = ?,
               ativo = ?,
               atualizado_em = ?,
               vencimento_certificado = ?
         WHERE id = ?;
        """,
        (nome.strip(), cnpj_num, novo_caminho, senha, nsu_int, ativo, agora, vencimento.strip(), empresa_id),
    )
    conn.commit()


def excluir_empresa(conn: sqlite3.Connection, empresa_id: int):
    cur = conn.cursor()
    cur.execute("DELETE FROM empresas WHERE id = ?;", (empresa_id,))
    conn.commit()


def set_ativo_para_todos(conn: sqlite3.Connection, ativo: int):
    cur = conn.cursor()
    cur.execute("UPDATE empresas SET ativo = ?;", (ativo,))
    conn.commit()


def set_ativo_para_ids(conn: sqlite3.Connection, ids: list[int], ativo: int):
    if not ids:
        return
    cur = conn.cursor()
    cur.executemany("UPDATE empresas SET ativo = ? WHERE id = ?;", [(ativo, i) for i in ids])
    conn.commit()


def get_empresa_por_id(conn: sqlite3.Connection, empresa_id: int):
    cur = conn.cursor()
    cur.execute(
        """
        SELECT id, nome, cnpj, caminho_certificado, senha_certificado,
               nsu_atual, ativo, criado_em, atualizado_em, vencimento_certificado
        FROM empresas
        WHERE id = ?;
        """,
        (empresa_id,),
    )
    return cur.fetchone()


# ==========================
# Certificados instalados (Windows)
# ==========================

def listar_certificados_instalados():
    """
    Usa PowerShell para listar certificados pessoais do usuário atual.
    Retorna lista de dicts: {FriendlyName, Subject, Thumbprint, NotAfterStr}
    onde NotAfterStr já vem formatado dd/MM/yyyy.
    """
    try:
        cmd = [
            "powershell",
            "-Command",
            (
                "Get-ChildItem Cert:\\CurrentUser\\My | "
                "Select-Object FriendlyName, Subject, Thumbprint, "
                "@{Name='NotAfterStr';Expression={$_.NotAfter.ToString('dd/MM/yyyy')}} | "
                "ConvertTo-Json"
            ),
        ]
        result = subprocess.run(
            cmd, capture_output=True, text=True, encoding="utf-8"
        )
        if result.returncode != 0 or not result.stdout.strip():
            return []

        data = json.loads(result.stdout)
        if isinstance(data, dict):
            data = [data]
        return data
    except Exception:
        return []


def exportar_certificado_instalado(thumbprint: str, senha: str, destino: str) -> bool:
    """
    Exporta um certificado do Windows (Cert:\CurrentUser\My) para um PFX.
    """
    try:
        cmd = [
            "powershell",
            "-Command",
            (
                f"$pwd = ConvertTo-SecureString '{senha}' -AsPlainText -Force; "
                f"Export-PfxCertificate -Cert \"Cert:\\CurrentUser\\My\\{thumbprint}\" "
                f"-FilePath \"{destino}\" -Password $pwd -Force"
            ),
        ]
        result = subprocess.run(
            cmd, capture_output=True, text=True, encoding="utf-8"
        )
        return result.returncode == 0
    except Exception:
        return False


# ==========================
# Popups
# ==========================

class CadastroNovoPopup(tk.Toplevel):
    """
    Popup para cadastro de nova empresa.
    Escolha entre:
      - Arquivo PFX/P12
      - Certificado instalado
    Depois retorna (caminho_pfx, senha) via callback.
    """
    def __init__(self, master, on_ok):
        super().__init__(master)
        self.title("Cadastrar nova empresa")
        self.resizable(False, False)
        self.on_ok = on_ok

        self.source_type = tk.StringVar(value="file")
        self.caminho_pfx = tk.StringVar()
        self.senha = tk.StringVar()
        self.selected_thumb = tk.StringVar()
        self.cert_list = []

        frm = tk.Frame(self, padx=10, pady=10)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Origem do certificado:").grid(row=0, column=0, sticky="w")
        rb_file = tk.Radiobutton(frm, text="Arquivo .pfx / .p12", variable=self.source_type, value="file", command=self._update_mode)
        rb_store = tk.Radiobutton(frm, text="Certificado instalado (Windows)", variable=self.source_type, value="store", command=self._update_mode)
        rb_file.grid(row=1, column=0, sticky="w")
        rb_store.grid(row=1, column=1, sticky="w")

        # Modo arquivo
        self.file_frame = tk.LabelFrame(frm, text="Arquivo PFX/P12", padx=5, pady=5)
        self.file_frame.grid(row=2, column=0, columnspan=2, sticky="we", pady=5)
        self.file_frame.columnconfigure(1, weight=1)

        tk.Label(self.file_frame, text="Arquivo:").grid(row=0, column=0, sticky="e")
        self.lbl_arquivo = tk.Label(self.file_frame, textvariable=self.caminho_pfx, width=40, anchor="w")
        self.lbl_arquivo.grid(row=0, column=1, sticky="we")
        btn_sel = tk.Button(self.file_frame, text="Selecionar...", command=self._selecionar_pfx)
        btn_sel.grid(row=0, column=2, padx=5)

        # Modo store
        self.store_frame = tk.LabelFrame(frm, text="Certificados instalados", padx=5, pady=5)
        self.store_frame.grid(row=3, column=0, columnspan=2, sticky="we", pady=5)
        self.store_frame.columnconfigure(0, weight=1)

        self.listbox = tk.Listbox(self.store_frame, width=70, height=6)
        self.listbox.grid(row=0, column=0, columnspan=2, sticky="we")
        btn_listar = tk.Button(self.store_frame, text="Atualizar lista", command=self._carregar_store)
        btn_listar.grid(row=1, column=0, sticky="w", pady=2)

        tk.Label(frm, text="Senha do certificado:").grid(row=4, column=0, sticky="e", pady=(5, 0))
        self.entry_senha = tk.Entry(frm, textvariable=self.senha, show="*", width=25)
        self.entry_senha.grid(row=4, column=1, sticky="w", pady=(5, 0))

        btns = tk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, pady=10)

        btn_ok = tk.Button(btns, text="OK", command=self._confirmar)
        btn_cancel = tk.Button(btns, text="Fechar", command=self.destroy)

        btn_ok.pack(side="left", padx=5)
        btn_cancel.pack(side="left", padx=5)


        self._update_mode()
        self._carregar_store()

        self.grab_set()
        self.transient(master)

    def _update_mode(self):
        mode = self.source_type.get()

        if mode == "file":
            for child in self.file_frame.winfo_children():
                try:
                    child.configure(state="normal")
                except tk.TclError:
                    pass
            for child in self.store_frame.winfo_children():
                try:
                    child.configure(state="disabled")
                except tk.TclError:
                    pass
        else:
            for child in self.file_frame.winfo_children():
                try:
                    child.configure(state="disabled")
                except tk.TclError:
                    pass
            for child in self.store_frame.winfo_children():
                try:
                    child.configure(state="normal")
                except tk.TclError:
                    pass



    def _selecionar_pfx(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar certificado (.pfx/.p12)",
            filetypes=[("Certificado PFX/P12", "*.pfx *.p12"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.caminho_pfx.set(caminho)

    def _carregar_store(self):
        self.listbox.delete(0, tk.END)
        self.cert_list = listar_certificados_instalados()
        if not self.cert_list:
            self.listbox.insert(tk.END, "Nenhum certificado encontrado ou erro ao listar.")
            return

        def format_cnpj(cnpj_digits: str) -> str:
            if len(cnpj_digits) != 14:
                return cnpj_digits
            return f"{cnpj_digits[:2]}.{cnpj_digits[2:5]}.{cnpj_digits[5:8]}/{cnpj_digits[8:12]}-{cnpj_digits[12:14]}"

        for c in self.cert_list:
            friendly = c.get("FriendlyName") or ""
            subject = c.get("Subject") or ""
            not_after = c.get("NotAfterStr") or ""
            thumb = c.get("Thumbprint") or ""

            # Tenta extrair CN do Subject se não tiver FriendlyName
            if not friendly and subject:
                m_cn = re.search(r"CN=([^,]+)", subject)
                if m_cn:
                    friendly = m_cn.group(1)

            if not friendly:
                friendly = "(sem nome)"

            # Tenta achar CNPJ (14 dígitos) no Subject
            cnpj_digits = ""
            m_cnpj = re.search(r"\d{14}", subject)
            if m_cnpj:
                cnpj_digits = m_cnpj.group(0)

            partes = [friendly]

            if cnpj_digits:
                partes.append(f"CNPJ {format_cnpj(cnpj_digits)}")

            if not_after:
                partes.append(f"vence {not_after}")

            if thumb:
                partes.append(f"thumb ...{thumb[-8:]}")

            linha = " | ".join(partes)
            self.listbox.insert(tk.END, linha)


    def _confirmar(self):
        mode = self.source_type.get()
        senha = self.senha.get() or ""
        if mode == "file":
            caminho = self.caminho_pfx.get().strip()
            if not caminho:
                messagebox.showwarning("Certificado", "Selecione um arquivo .pfx/.p12.")
                return
            if not os.path.isfile(caminho):
                messagebox.showerror("Arquivo", f"Arquivo não encontrado:\n{caminho}")
                return
            self.on_ok("file", caminho, senha, None)
            self.destroy()
        else:
            if not self.cert_list:
                messagebox.showerror("Certificados", "Nenhum certificado instalado foi encontrado.")
                return
            sel = self.listbox.curselection()
            if not sel:
                messagebox.showwarning("Certificados", "Selecione um certificado instalado.")
                return
            escolhido = self.cert_list[sel[0]]
            thumb = escolhido.get("Thumbprint")
            if not thumb:
                messagebox.showerror("Certificados", "Não foi possível obter o Thumbprint.")
                return
            self.on_ok("store", None, senha, thumb)
            self.destroy()


class EditarEmpresaPopup(tk.Toplevel):
    """
    Popup para edição completa de uma empresa.
    """
    def __init__(self, master, empresa_row, on_save):
        super().__init__(master)
        self.title("Editar empresa")
        self.resizable(False, False)
        self.on_save = on_save

        (
            self.empresa_id,
            nome,
            cnpj,
            caminho,
            senha,
            nsu,
            ativo,
            criado,
            atualizado,
            vencimento,
        ) = empresa_row

        frame = tk.Frame(self, padx=10, pady=10)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Nome:").grid(row=0, column=0, sticky="e")
        tk.Label(frame, text="CNPJ:").grid(row=1, column=0, sticky="e")
        tk.Label(frame, text="Caminho certificado:").grid(row=2, column=0, sticky="e")
        tk.Label(frame, text="Senha certificado:").grid(row=3, column=0, sticky="e")
        tk.Label(frame, text="Vencimento (dd/mm/aaaa):").grid(row=4, column=0, sticky="e")
        tk.Label(frame, text="NSU atual:").grid(row=5, column=0, sticky="e")
        tk.Label(frame, text="Ativo:").grid(row=6, column=0, sticky="e")

        self.entry_nome = tk.Entry(frame, width=50)
        self.entry_cnpj = tk.Entry(frame, width=20)
        self.entry_caminho = tk.Entry(frame, width=50)
        self.entry_senha = tk.Entry(frame, width=20, show="*")
        self.entry_venc = tk.Entry(frame, width=15)
        self.entry_nsu = tk.Entry(frame, width=15)
        self.var_ativo = tk.BooleanVar(value=bool(ativo))

        self.entry_nome.grid(row=0, column=1, columnspan=2, sticky="w")
        self.entry_cnpj.grid(row=1, column=1, sticky="w")
        self.entry_caminho.grid(row=2, column=1, sticky="w")
        self.entry_senha.grid(row=3, column=1, sticky="w")
        self.entry_venc.grid(row=4, column=1, sticky="w")
        self.entry_nsu.grid(row=5, column=1, sticky="w")

        self.chk_ativo = tk.Checkbutton(frame, text="Ativo para busca", variable=self.var_ativo)
        self.chk_ativo.grid(row=6, column=1, sticky="w")

        btn_sel = tk.Button(frame, text="Selecionar novo certificado...", command=self._selecionar_pfx)
        btn_sel.grid(row=2, column=2, padx=5, sticky="w")

        self.entry_nome.insert(0, nome or "")
        self.entry_cnpj.insert(0, cnpj or "")
        self.entry_caminho.insert(0, caminho or "")
        # senha não é recarregada por segurança
        if vencimento:
            self.entry_venc.insert(0, vencimento)
        self.entry_nsu.insert(0, str(nsu or 0))

        btns = tk.Frame(frame)
        btns.grid(row=7, column=0, columnspan=3, pady=10)
        tk.Button(btns, text="Cancelar", command=self.destroy).pack(side="right", padx=5)
        tk.Button(btns, text="Salvar", command=self._salvar).pack(side="right", padx=5)

        self.grab_set()
        self.transient(master)

    def _selecionar_pfx(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar certificado (.pfx/.p12)",
            filetypes=[("Certificado PFX/P12", "*.pfx *.p12"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.entry_caminho.delete(0, tk.END)
            self.entry_caminho.insert(0, caminho)

    def _salvar(self):
        nome = self.entry_nome.get().strip()
        cnpj = self.entry_cnpj.get().strip()
        caminho = self.entry_caminho.get().strip()
        senha = self.entry_senha.get()
        venc = self.entry_venc.get().strip()
        nsu = self.entry_nsu.get().strip()
        ativo = 1 if self.var_ativo.get() else 0

        if not nome or not cnpj:
            messagebox.showwarning("Dados insuficientes", "Preencha pelo menos Nome e CNPJ.")
            return

        self.on_save(self.empresa_id, nome, cnpj, caminho, senha, nsu, venc, ativo)
        self.destroy()


# ==========================
# GUI principal
# ==========================

class GerenciadorEmpresasApp:
    def __init__(self, master, conn: sqlite3.Connection):
        self.master = master
        self.conn = conn
        master.title("Cadastro de Empresas - NFe Distribuição DF-e")

        self.current_filter = ""
        self.sort_reverse = {}

        # ---------- Topo: botão de cadastro + botões de execução ----------
        top_frame = tk.Frame(master)
        top_frame.pack(fill="x", padx=10, pady=5)

        # Botão cadastrar nova empresa
        self.btn_cadastrar = tk.Button(top_frame, text="Cadastrar nova empresa", command=self.abrir_cadastro_novo)
        self.btn_cadastrar.pack(side="left")

        exec_frame = tk.Frame(top_frame)
        exec_frame.pack(side="right", anchor="ne", padx=10)

        self.btn_exec_sel = tk.Button(
            exec_frame,
            text="Salvar & executar selecionadas",
            command=self.salvar_e_executar_selecionadas,
            width=22
        )
        self.btn_exec_all = tk.Button(
            exec_frame,
            text="Salvar & executar ativas",
            command=self.salvar_e_executar_todas,
            width=22
        )
        self.btn_sair = tk.Button(
            exec_frame,
            text="Fechar (sem executar)",
            command=self.master.destroy,
            width=18
        )

        # agora os 3 botões ficam lado a lado
        self.btn_exec_sel.pack(side="left", padx=5)
        self.btn_exec_all.pack(side="left", padx=5)
        self.btn_sair.pack(side="left", padx=5)


        # ---------- Busca ----------
        busca_frame = tk.Frame(master)
        busca_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(busca_frame, text="Buscar (Nome/CNPJ):").pack(side="left")
        self.entry_busca = tk.Entry(busca_frame, width=30)
        self.entry_busca.pack(side="left", padx=5)

        btn_buscar = tk.Button(busca_frame, text="Buscar", command=self.aplicar_filtro)
        btn_limpar = tk.Button(busca_frame, text="Limpar", command=self.limpar_filtro)
        btn_buscar.pack(side="left")
        btn_limpar.pack(side="left", padx=5)

        # ---------- Lista ----------
        lista_frame = tk.LabelFrame(master, text="Empresas cadastradas", padx=10, pady=10)
        lista_frame.pack(fill="both", expand=True, padx=10, pady=5)

        colunas = ("nome", "cnpj", "vencimento", "nsu", "ativo")
        self.tree = ttk.Treeview(
            lista_frame,
            columns=colunas,
            show="headings",
            height=10,
            selectmode="extended",
        )
        self.tree.heading("nome", text="Nome", command=lambda c="nome": self.ordenar_por(c))
        self.tree.heading("cnpj", text="CNPJ", command=lambda c="cnpj": self.ordenar_por(c))
        self.tree.heading("vencimento", text="Vencimento", command=lambda c="vencimento": self.ordenar_por(c))
        self.tree.heading("nsu", text="NSU Atual", command=lambda c="nsu": self.ordenar_por(c))
        self.tree.heading("ativo", text="Ativo", command=lambda c="ativo": self.ordenar_por(c))

        self.tree.column("nome", width=240)
        self.tree.column("cnpj", width=120)
        self.tree.column("vencimento", width=110, anchor="center")
        self.tree.column("nsu", width=80, anchor="e")
        self.tree.column("ativo", width=60, anchor="center")

        self.tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(lista_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # Tags
        self.tree.tag_configure("ativo", foreground="green")
        self.tree.tag_configure("inativo", foreground="red")
        self.tree.tag_configure("vencido", background="#ffcccc")
        self.tree.tag_configure("venc_alarme", background="#ffe0b3")

        # Eventos
        self.tree.bind("<<TreeviewSelect>>", self.on_select_row)
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_double_click)

        # ---------- Botões inferiores ----------
        btn_frame = tk.Frame(master)
        btn_frame.pack(fill="x", padx=10, pady=5)

        btn_refresh = tk.Button(btn_frame, text="Atualizar lista", command=self.carregar_lista)
        btn_refresh.pack(side="left")

        btn_marcar_todos = tk.Button(btn_frame, text="Ativar todos", command=self.marcar_todos)
        btn_desmarcar_todos = tk.Button(btn_frame, text="Desativar todos", command=self.desmarcar_todos)
        btn_marcar_todos.pack(side="left", padx=5)
        btn_desmarcar_todos.pack(side="left", padx=5)

        self.sel_btn_frame = tk.Frame(btn_frame)
        self.btn_del = tk.Button(self.sel_btn_frame, text="Excluir selecionadas", command=self.excluir_selecionadas)
        self.btn_ativar_sel = tk.Button(self.sel_btn_frame, text="Ativar selecionadas", command=self.ativar_selecionadas)
        self.btn_desativar_sel = tk.Button(self.sel_btn_frame, text="Desativar selecionadas", command=self.desativar_selecionadas)
        self.btn_editar_sel = tk.Button(self.sel_btn_frame, text="Editar selecionada", command=self.editar_selecionada)

        self.btn_del.pack(side="left", padx=5)
        self.btn_ativar_sel.pack(side="left", padx=5)
        self.btn_desativar_sel.pack(side="left", padx=5)
        self.btn_editar_sel.pack(side="left", padx=5)

        self.sel_btn_frame.pack_forget()

        self.carregar_lista()

    # ---------- Cadastro novo ----------

    def abrir_cadastro_novo(self):
        def on_ok(mode, caminho, senha, thumb):
            if mode == "file":
                self._cadastrar_por_arquivo(caminho, senha)
            else:
                self._cadastrar_por_cert_instalado(thumb, senha)

        CadastroNovoPopup(self.master, on_ok)

    def _cadastrar_por_arquivo(self, caminho, senha):
        if not CRYPTO_OK:
            messagebox.showerror(
                "cryptography não instalada",
                "Para importar automaticamente dados do certificado, instale a biblioteca:\n\npip install cryptography"
            )
            return

        nome, cnpj, venc = ler_dados_certificado_completo(caminho, senha)
        if not nome or not cnpj:
            messagebox.showerror(
                "Falha ao ler certificado",
                "Não foi possível ler Nome/CNPJ do certificado.\n"
                "Verifique a senha do PFX e o tipo de certificado."
            )
            return

        try:
            inserir_empresa(self.conn, nome, cnpj, caminho, senha, "0", venc or "")
        except sqlite3.IntegrityError as e:
            messagebox.showerror("Erro ao cadastrar", f"Erro de integridade (CNPJ duplicado?):\n{e}")
            return
        except Exception as e:
            messagebox.showerror("Erro ao cadastrar", f"Ocorreu um erro:\n{e}")
            return

        messagebox.showinfo("Sucesso", "Empresa cadastrada com sucesso!")
        self.carregar_lista()

    def _cadastrar_por_cert_instalado(self, thumb, senha):
        if not CRYPTO_OK:
            messagebox.showerror(
                "cryptography não instalada",
                "Para importar automaticamente dados do certificado, instale a biblioteca:\n\npip install cryptography"
            )
            return

        base_dir = get_script_dir()
        temp_pfx = os.path.join(base_dir, f"_temp_{thumb}.pfx")
        ok = exportar_certificado_instalado(thumb, senha, temp_pfx)
        if not ok or not os.path.isfile(temp_pfx):
            messagebox.showerror(
                "Exportação",
                "Não foi possível exportar o certificado instalado.\n"
                "Verifique a senha e se o certificado permite exportação."
            )
            return

        nome, cnpj, venc = ler_dados_certificado_completo(temp_pfx, senha)
        if not nome or not cnpj:
            os.remove(temp_pfx)
            messagebox.showerror(
                "Falha ao ler certificado",
                "Não foi possível ler Nome/CNPJ do certificado exportado.\n"
                "Verifique o certificado."
            )
            return

        try:
            inserir_empresa(self.conn, nome, cnpj, temp_pfx, senha, "0", venc or "")
        except sqlite3.IntegrityError as e:
            os.remove(temp_pfx)
            messagebox.showerror("Erro ao cadastrar", f"Erro de integridade (CNPJ duplicado?):\n{e}")
            return
        except Exception as e:
            os.remove(temp_pfx)
            messagebox.showerror("Erro ao cadastrar", f"Ocorreu um erro:\n{e}")
            return

        try:
            os.remove(temp_pfx)
        except Exception:
            pass

        messagebox.showinfo("Sucesso", "Empresa cadastrada com sucesso (certificado instalado)!")
        self.carregar_lista()

    # ---------- Lista / seleções ----------

    def _tags_vencimento(self, vencimento: str):
        vencimento = (vencimento or "").strip()
        if not vencimento:
            return None
        try:
            dt_venc = datetime.strptime(vencimento, "%d/%m/%Y").date()
        except ValueError:
            return None

        hoje = date.today()
        if dt_venc < hoje:
            return "vencido"
        dias = (dt_venc - hoje).days
        if dias <= 14:
            return "venc_alarme"
        return None

    def carregar_lista(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        rows = carregar_empresas(self.conn, self.current_filter)
        for r in rows:
            (
                id_,
                nome,
                cnpj,
                caminho,
                senha,
                nsu,
                ativo,
                criado,
                atualizado,
                vencimento,
            ) = r

            checkbox = "✔" if ativo else "✖"
            tag_ativo = "ativo" if ativo else "inativo"
            tag_venc = self._tags_vencimento(vencimento)
            tags = [tag_ativo]
            if tag_venc:
                tags.append(tag_venc)

            self.tree.insert(
                "",
                tk.END,
                iid=str(id_),
                values=(nome, cnpj, vencimento or "", nsu, checkbox),
                tags=tuple(tags),
            )

        self.atualizar_botoes_selecao()

    def excluir_selecionadas(self):
        sel = self.tree.selection()
        if not sel:
            return

        if not messagebox.askyesno("Confirmar exclusão", "Excluir todas as empresas selecionadas?"):
            return

        for item in sel:
            empresa_id = int(item)
            try:
                excluir_empresa(self.conn, empresa_id)
            except Exception as e:
                messagebox.showerror("Erro ao excluir", f"Ocorreu um erro ao excluir ID {empresa_id}:\n{e}")
                return

        messagebox.showinfo("Sucesso", "Empresas excluídas.")
        self.carregar_lista()

    def on_select_row(self, event):
        self.atualizar_botoes_selecao()

    def on_double_click(self, event):
        rowid = self.tree.identify_row(event.y)
        if not rowid:
            return
        self.abrir_edicao(int(rowid))

    def abrir_edicao(self, empresa_id: int):
        row = get_empresa_por_id(self.conn, empresa_id)
        if not row:
            return

        def salvar_edicao(empresa_id, nome, cnpj, caminho, senha, nsu, venc, ativo):
            try:
                atualizar_empresa(self.conn, empresa_id, nome, cnpj, caminho, senha, nsu, venc, ativo)
            except Exception as e:
                messagebox.showerror("Erro ao salvar edição", f"Ocorreu um erro:\n{e}")
                return
            self.carregar_lista()

        EditarEmpresaPopup(self.master, row, salvar_edicao)

    # ---------- Busca ----------

    def aplicar_filtro(self):
        self.current_filter = self.entry_busca.get().strip()
        self.carregar_lista()

    def limpar_filtro(self):
        self.current_filter = ""
        self.entry_busca.delete(0, tk.END)
        self.carregar_lista()

    # ---------- Ordenação ----------

    def ordenar_por(self, coluna: str):
        data = []
        for item in self.tree.get_children(""):
            valor = self.tree.set(item, coluna)
            data.append((valor, item))

        def try_cast(v):
            try:
                return int(v)
            except ValueError:
                return v

        rev = self.sort_reverse.get(coluna, False)
        self.sort_reverse[coluna] = not rev

        data.sort(key=lambda t: try_cast(t[0]), reverse=rev)
        for index, (_, item) in enumerate(data):
            self.tree.move(item, "", index)

    # ---------- Ativo / Inativo ----------

    def marcar_todos(self):
        try:
            set_ativo_para_todos(self.conn, 1)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ativar todas as empresas:\n{e}")
            return
        self.carregar_lista()

    def desmarcar_todos(self):
        try:
            set_ativo_para_todos(self.conn, 0)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao desativar todas as empresas:\n{e}")
            return
        self.carregar_lista()

    def ativar_selecionadas(self):
        sel = self.tree.selection()
        if not sel:
            return
        ids = [int(item) for item in sel]
        try:
            set_ativo_para_ids(self.conn, ids, 1)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ativar selecionadas:\n{e}")
            return
        self.carregar_lista()

    def desativar_selecionadas(self):
        sel = self.tree.selection()
        if not sel:
            return
        ids = [int(item) for item in sel]
        try:
            set_ativo_para_ids(self.conn, ids, 0)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao desativar selecionadas:\n{e}")
            return
        self.carregar_lista()

    def atualizar_botoes_selecao(self):
        sel = self.tree.selection()
        if sel and not self.sel_btn_frame.winfo_ismapped():
            self.sel_btn_frame.pack(side="left", padx=10)
        elif not sel and self.sel_btn_frame.winfo_ismapped():
            self.sel_btn_frame.pack_forget()

    def editar_selecionada(self):
        sel = self.tree.selection()
        if not sel:
            return
        self.abrir_edicao(int(sel[0]))

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        col = self.tree.identify_column(event.x)
        rowid = self.tree.identify_row(event.y)
        if not rowid:
            return

        # coluna 'ativo'
        if col == "#5":
            empresa_id = int(rowid)
            current = self.tree.set(rowid, "ativo")
            novo_ativo = 0 if current == "✔" else 1

            try:
                set_ativo_para_ids(self.conn, [empresa_id], novo_ativo)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao alterar ativo/inativo:\n{e}")
                return "break"

            checkbox = "✔" if novo_ativo else "✖"
            tag_ativo = "ativo" if novo_ativo else "inativo"

            vals = list(self.tree.item(rowid, "values"))
            vencimento = vals[2]
            tag_venc = self._tags_vencimento(vencimento)
            tags = [tag_ativo]
            if tag_venc:
                tags.append(tag_venc)

            vals[-1] = checkbox
            self.tree.item(rowid, values=vals, tags=tuple(tags))

            self.atualizar_botoes_selecao()
            return "break"

    # ---------- Execução do script principal ----------

    def salvar_e_executar_todas(self):
        self._executar_script_busca()

    def salvar_e_executar_selecionadas(self):
        sel = self.tree.selection()
        if not sel:
            if not messagebox.askyesno(
                "Nenhuma seleção",
                "Nenhuma empresa selecionada.\nDeseja executar para TODAS as empresas ativas?"
            ):
                return
            self._executar_script_busca()
            return

        ids_selecionados = [int(item) for item in sel]

        try:
            set_ativo_para_todos(self.conn, 0)
            set_ativo_para_ids(self.conn, ids_selecionados, 1)
        except Exception as e:
            messagebox.showerror("Erro ao preparar execução", f"Erro ao atualizar flag 'ativo':\n{e}")
            return

        self._executar_script_busca()

    def _executar_script_busca(self):
        base_dir = get_script_dir()
        script_path = os.path.join(base_dir, "projeto fsist v1.py")

        if not os.path.isfile(script_path):
            messagebox.showerror(
                "Script não encontrado",
                f"Não encontrei o script de busca:\n{script_path}"
            )
            return

        self.master.destroy()
        try:
            subprocess.Popen([sys.executable, script_path])
        except Exception as e:
            messagebox.showerror("Erro ao executar busca", f"Erro ao chamar o script de busca:\n{e}")


# ==========================
# main
# ==========================

def main():
    base_dir = get_script_dir()
    db_path = os.path.join(base_dir, "certificados.db")

    conn = sqlite3.connect(db_path)
    init_db(conn)

    root = tk.Tk()
    root.geometry("900x600")
    root.resizable(False, False)

    app = GerenciadorEmpresasApp(root, conn)
    root.mainloop()

    conn.close()


if __name__ == "__main__":
    main()
