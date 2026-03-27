# setup_cert_profile.py
from playwright.sync_api import sync_playwright
import os
import tempfile
import time

# Caminho do perfil que será reaproveitado depois (altere para a sua pasta)
PROFILE_DIR = os.path.join(tempfile.gettempdir(), "perfil_firefox_cert_esnfs")
# Exemplo Windows: PROFILE_DIR = r"C:\Users\SeuUsuario\firefox_profile_esnfs"

URL = "https://www.esnfs.com.br/?e=35"

def main():
    os.makedirs(PROFILE_DIR, exist_ok=True)
    print(f"Usando perfil em: {PROFILE_DIR}")
    print("Abrindo Firefox. Se aparecer a janela de certificado, escolha o certificado que deseja e marque 'Lembrar sempre' (ou similar).")
    print("Quando terminar, feche o navegador para salvar o perfil e execute o buscador que utilizará este perfil.")

    with sync_playwright() as p:
        # Cria contexto persistente para que as escolhas (cookies, preferências) sejam salvas no perfil
        contexto = p.firefox.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=False,
            accept_downloads=True,
            # não forçamos seleção automática aqui — queremos que você selecione manualmente
        )
        pagina = contexto.new_page()
        pagina.goto(URL)
        pagina.wait_for_load_state("domcontentloaded")
        # Dá tempo pra você interagir manualmente (selecionar certificado)
        try:
            # espera longa para que o usuário faça a seleção (ajuste se quiser)
            for i in range(600):
                time.sleep(1)
                # Opcional: exibe se houver iframe de recaptcha (só info)
                if pagina.locator("iframe[title*='recaptcha']").count() > 0:
                    print("Detectado iframe de reCAPTCHA na página.")
        except KeyboardInterrupt:
            print("Interrupção manual... fechando.")
        finally:
            try:
                contexto.close()
            except Exception:
                pass
            print("Perfil salvo. Agora você pode usar esse PROFILE_DIR no seu script principal.")

if __name__ == "__main__":
    main()
