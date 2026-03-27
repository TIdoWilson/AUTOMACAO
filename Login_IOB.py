from pywinauto import Application, mouse, Desktop
import time
import argparse
import os
import getpass

LOGIN_FATOR_LARGURA = 0.4791
LOGIN_FATOR_ALTURA  = 0.6328

CAMINHO_EXE = r"C:\IOBGestaoContabil\Programas\cache.exe"

def abrir_programa():
    print(f"Abrindo o programa: {CAMINHO_EXE}")
    app = Application(backend="uia").start(CAMINHO_EXE)
    # PID do processo aberto
    pid = app.process
    return app, pid

def esperar_tela_login(timeout=60):
    print("Aguardando a tela 'Login SSO' aparecer...")
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            dlg = Desktop(backend="uia").window(title_re=".*Login.*SSO.*")
            if dlg.exists() and dlg.is_visible():
                print("✔ Tela de Login encontrada!")
                return dlg
        except Exception:
            pass
        time.sleep(0.5)
    raise RuntimeError("❌ A tela 'Login SSO' não apareceu dentro do tempo limite.")

def achar_campos_login(dlg):
    edits = dlg.descendants(control_type="Edit")
    if len(edits) < 2:
        raise RuntimeError(f"Esperava pelo menos 2 campos Edit, encontrei {len(edits)}")

    campo_usuario = None
    campo_senha = None

    for edit in edits:
        try:
            if edit.is_password():
                campo_senha = edit
            else:
                if campo_usuario is None:
                    campo_usuario = edit
        except Exception:
            continue

    if campo_usuario is None or campo_senha is None:
        campo_usuario = edits[0]
        campo_senha = edits[1]

    return campo_usuario, campo_senha

def fazer_login(dlg_login, usuario, senha):
    print("Localizando campos de usuário e senha...")
    campo_usuario, campo_senha = achar_campos_login(dlg_login)

    pnl = campo_usuario.element_info.parent
    rect = pnl.rectangle
    largura = rect.right - rect.left
    altura = rect.bottom - rect.top

    print("Preenchendo usuário...")
    campo_usuario.set_focus()
    campo_usuario.type_keys("^a{BACKSPACE}")
    campo_usuario.type_keys(usuario, with_spaces=True)

    print("Preenchendo senha...")
    campo_senha.set_focus()
    campo_senha.type_keys("^a{BACKSPACE}")
    campo_senha.type_keys(senha, with_spaces=True)

    time.sleep(0.5)

    x = int(rect.left + largura * LOGIN_FATOR_LARGURA)
    y = int(rect.top + altura * LOGIN_FATOR_ALTURA)

    print(f"Clicando em 'Acessar com login tradicional' em ({x}, {y})...")
    mouse.click(button="left", coords=(x, y))

    print("Aguardando pós-login...")
    time.sleep(2)

def fechar_programa(pid, timeout=10):
    """
    Tenta fechar de forma limpa; se não der, mata o processo.
    """
    print("Fechando o programa...")

    try:
        app = Application(backend="uia").connect(process=pid)
        w = app.top_window()
        # fecha "normal"
        w.close()
        # espera sumir
        w.wait_not("visible", timeout=timeout)
        print("✔ Programa fechado normalmente.")
        return
    except Exception as e:
        print(f"⚠ Não fechou normalmente ({e}). Forçando encerramento...")

    try:
        app = Application(backend="uia").connect(process=pid)
        app.kill()
        print("✔ Processo encerrado (kill).")
    except Exception as e:
        print(f"❌ Falha ao matar o processo: {e}")

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--usuario", "-u")
    p.add_argument("--senha", "-p")
    p.add_argument("--timeout", type=int, default=60)
    p.add_argument("--fechar", action="store_true", help="Fecha o IOB ao concluir.")
    return p.parse_args()

def main():
    args = parse_args()

    usuario = args.usuario or os.environ.get("IOB_USUARIO") or input("Usuário: ").strip()
    senha = args.senha or os.environ.get("IOB_SENHA") or getpass.getpass("Senha: ")

    app, pid = abrir_programa()
    dlg_login = esperar_tela_login(timeout=args.timeout)
    fazer_login(dlg_login, usuario, senha)

    # Se quiser fechar automaticamente no final:
    if args.fechar:
        fechar_programa(pid)

if __name__ == "__main__":
    main()
