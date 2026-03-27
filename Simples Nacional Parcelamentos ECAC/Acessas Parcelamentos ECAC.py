import os
import random
import time
import traceback
from dotenv import load_dotenv
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Importa o launcher corrigido
from chrome_5543_launcher import abrir_chrome_5543, ECAC_PROFILE_PATH
from hcaptcha_utils import resolver_hcaptcha

load_dotenv()

# ========== CONFIGURAÇÕES DO .ENV ==========
CNPJS_LIST = os.getenv("CNPJS_LIST", "")
CNPJ_FILE_PATH = os.getenv("CNPJ_FILE_PATH", "")
OPEN_IN_NEW_TAB = os.getenv("OPEN_IN_NEW_TAB", "1").strip() not in {"0", "false", "False"}
MODO_LOGIN_MANUAL = os.getenv("MODO_LOGIN_MANUAL", "1").strip() not in {"0", "false", "False"}

ECAC_URLS = [
    "https://cav.receita.fazenda.gov.br/autenticacao/Login/IndexGovBr",
    "https://cav.receita.fazenda.gov.br/autenticacao",
    "https://cav.receita.fazenda.gov.br/ecac/",
]
# =========================================


# ========== FUNÇÕES DE COMPORTAMENTO HUMANO ==========
def esperar(min_s=0.5, max_s=1.5):
    time.sleep(random.uniform(min_s, max_s))


def mover_mouse_aleatorio(driver):
    largura = driver.execute_script("return window.innerWidth;")
    altura = driver.execute_script("return window.innerHeight;")
    x = random.randint(100, largura - 100)
    y = random.randint(100, altura - 100)
    ActionChains(driver).move_by_offset(x, y).perform()
    ActionChains(driver).move_by_offset(-x, -y).perform()


def rolar_aleatorio(driver):
    altura = driver.execute_script("return document.body.scrollHeight")
    if altura > 0:
        pos = random.randint(0, altura)
        driver.execute_script(f"window.scrollTo(0, {pos});")
        esperar(0.2, 0.6)


def clicar_humano(elemento, driver, mover_antes=True):
    actions = ActionChains(driver)
    if mover_antes:
        actions.move_to_element(elemento).perform()
        esperar(0.1, 0.3)
    actions.click(elemento).perform()


def digitar_como_humano(campo, texto):
    for caractere in texto:
        campo.send_keys(caractere)
        time.sleep(random.uniform(0.05, 0.12))
# =========================================


def carregar_cnpjs():
    if CNPJS_LIST:
        return [cnpj.strip() for cnpj in CNPJS_LIST.split(",") if cnpj.strip()]
    if CNPJ_FILE_PATH and os.path.exists(CNPJ_FILE_PATH):
        with open(CNPJ_FILE_PATH, "r", encoding="utf-8") as arquivo:
            return [linha.strip() for linha in arquivo if linha.strip()]
    return ["12345678000199"]


def abrir_fluxo_ecac(driver):
    print("🔄 Iniciando abertura do eCAC...")
    driver.get("about:blank")
    if OPEN_IN_NEW_TAB:
        print("🔄 Abrindo fluxo em nova aba")
        driver.switch_to.new_window("tab")

    for url in ECAC_URLS:
        for tentativa in range(1, 4):
            print(f"🔄 Tentando abrir: {url} | tentativa {tentativa}")
            try:
                driver.get(url)
            except TimeoutException:
                print("⏱️ Timeout de carregamento. Interrompendo página e tentando novamente.")
                driver.execute_script("window.stop();")

            esperar(1.2, 2.0)
            if "cav.receita.fazenda.gov.br" in (driver.current_url or "").lower():
                print(f"✅ Site aberto: {driver.current_url}")
                return
            print(f"⚠️ Tentativa {tentativa} sem sucesso para {url}. URL atual: {driver.current_url}")

    raise TimeoutException(f"Não conseguiu abrir o eCAC. URL atual: {driver.current_url}")


def executar_login_automatico(driver, wait):
    mover_mouse_aleatorio(driver)
    esperar(0.5, 1.0)

    btn_gov = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@alt='Acesso Gov BR']")))
    clicar_humano(btn_gov, driver)
    print("✅ Botão Gov BR clicado")

    prazo = time.time() + 90
    while time.time() < prazo:
        if driver.find_elements(By.ID, "btnPerfil"):
            print("✅ eCAC autenticado e botão de perfil disponível")
            return
        if driver.find_elements(By.ID, "login-certificate"):
            btn_cert = wait.until(EC.element_to_be_clickable((By.ID, "login-certificate")))
            clicar_humano(btn_cert, driver)
            print("✅ 'Seu certificado digital' selecionado")
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "btnPerfil")))
            print("✅ Recarregou para o modal de perfis após o certificado")
            return
        esperar(0.3, 0.8)

    raise TimeoutException("Timeout no login GovBR.")


def aguardar_login_manual(wait):
    print("🔐 Modo login manual ativado.")
    print("🔐 Resolva login/certificado/captcha no navegador e pressione ENTER no terminal para continuar.")
    input("Pressione ENTER para continuar...")
    wait.until(EC.presence_of_element_located((By.ID, "btnPerfil")))
    print("✅ Sessão validada. Cookies reaproveitados no perfil exclusivo.")


def abrir_modal_perfil(driver, wait):
    rolar_aleatorio(driver)
    esperar(0.5, 1.2)

    btn_perfil = wait.until(EC.element_to_be_clickable((By.ID, "btnPerfil")))
    clicar_humano(btn_perfil, driver)
    print("✅ 'Alterar perfil de acesso' clicado")
    esperar(2, 3)


def preencher_cnpj_e_alterar(driver, wait, cnpj):
    cnpj_input = wait.until(EC.presence_of_element_located((By.ID, "txtNIPapel2")))
    cnpj_input.clear()
    esperar(0.3, 0.8)
    digitar_como_humano(cnpj_input, cnpj)
    print(f"✅ CNPJ {cnpj} preenchido")

    total_tempo_espera = 15.0
    espera_inicial = random.uniform(1.5, 2.5)
    print(f"⏳ Aguardando {espera_inicial:.1f}s antes de deletar os últimos 3 caracteres...")
    time.sleep(espera_inicial)

    for _ in range(3):
        cnpj_input.send_keys(Keys.BACKSPACE)
    print("🗑️ Últimos 3 caracteres deletados")

    ultimos_tres = cnpj[-3:]
    print(f"✍️ Redigitando '{ultimos_tres}' lentamente...")
    for caractere in ultimos_tres:
        cnpj_input.send_keys(caractere)
        time.sleep(random.uniform(0.05, 0.12))
    print("✅ Últimos 3 caracteres redigitados")

    tempo_passado = espera_inicial + (3 * 0.08)
    tempo_restante = max(0.5, total_tempo_espera - tempo_passado)
    print(f"⏳ Aguardando mais {tempo_restante:.1f}s para completar os 15 segundos...")
    time.sleep(tempo_restante)

    cnpj_input.send_keys(Keys.ENTER)
    print("✅ Enter pressionado (confirmação)")


def resolver_hcaptcha_perfil(driver, timeout):
    if not resolver_hcaptcha(driver, timeout=timeout):
        print("⚠️ Captcha não resolvido via iframe – verifique se o captcha foi resolvido manualmente.")
        return False
    else:
        print("✅ Captcha resolvido")
        esperar(2, 3)
        return True


def navegar_ate_emissao_parcela(driver, wait):
    rolar_aleatorio(driver)
    esperar(0.5, 1.2)

    btn_simples = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Simples Nacional')]")))
    clicar_humano(btn_simples, driver)
    print("✅ Menu 'Simples Nacional' acessado")
    esperar(2, 3)

    rolar_aleatorio(driver)
    link_das = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Solicitar, acompanhar e emitir DAS de parcelamento')]"))
    )
    clicar_humano(link_das, driver)
    print("✅ Link de DAS de parcelamento clicado")
    esperar(2, 3)

    rolar_aleatorio(driver)
    btn_emissao = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_contentPlaceH_emitirDas")))
    clicar_humano(btn_emissao, driver)
    print("✅ Botão 'Emissão de Parcela' acionado")


def retornar_para_tela_perfil(driver):
    driver.back()
    esperar(1, 2)
    driver.back()
    esperar(1, 2)
    driver.back()
    esperar(2, 3)


def processar_cnpj(driver, wait, cnpj, index, total):
    print(f"\n===== ({index}/{total}) Processando CNPJ: {cnpj} =====")
    max_retries = 3
    for tentativa in range(1, max_retries + 1):
        if tentativa > 1:
            print(f"🔄 Recarregando página e tentando novamente (tentativa {tentativa})...")
            driver.refresh()
            esperar(3, 5)
            # Re-abrir modal de perfil após reload
            abrir_modal_perfil(driver, wait)

        try:
            preencher_cnpj_e_alterar(driver, wait, cnpj)
            if resolver_hcaptcha_perfil(driver, timeout=20):
                # Captcha resolvido, então sai do loop de tentativas
                break
            else:
                # Captcha não resolvido, continua para próxima tentativa
                print(f"⚠️ Captcha não resolvido na tentativa {tentativa}. Tentando novamente...")
                continue
        except Exception as e:
            print(f"⚠️ Falha na tentativa {tentativa}: {e}")
            if tentativa == max_retries:
                print(f"❌ Não foi possível alterar o perfil para CNPJ {cnpj} após {max_retries} tentativas. Pulando...")
                return
            continue

    # Se chegou aqui, conseguiu alterar o perfil, agora navega até emissão
    try:
        navegar_ate_emissao_parcela(driver, wait)
        resolver_hcaptcha_perfil(driver, timeout=25)
        print(f"✅ Fluxo atingiu a etapa 'Emissão de Parcela' para {cnpj}.")
    except Exception as e:
        print(f"⚠️ Falha durante a navegação para emissão de parcela: {e}. Continuando para próximo CNPJ.")

    retornar_para_tela_perfil(driver)


def executar_fluxo_principal(driver, wait, cnpjs):
    abrir_fluxo_ecac(driver)
    esperar(2, 3)

    if MODO_LOGIN_MANUAL:
        aguardar_login_manual(wait)
    else:
        executar_login_automatico(driver, wait)

    abrir_modal_perfil(driver, wait)

    total = len(cnpjs)
    for index, cnpj in enumerate(cnpjs, start=1):
        processar_cnpj(driver, wait, cnpj, index, total)


def main():
    cnpjs = carregar_cnpjs()
    if not cnpjs:
        print("❌ Nenhum CNPJ encontrado. Ajuste a configuração.")
        return

    driver = None
    try:
        print("✅ Criando sessão do navegador com undetected-chromedriver...")
        print(f"✅ Perfil dedicado em uso: {ECAC_PROFILE_PATH}")
        driver = abrir_chrome_5543()
        wait = WebDriverWait(driver, 20)
        executar_fluxo_principal(driver, wait, cnpjs)

    except Exception as erro:
        print(f"❌ Erro geral: {erro}")
        print(traceback.format_exc())
        if driver:
            driver.save_screenshot("erro_parcelamento.png")
            print("Screenshot salva como erro_parcelamento.png")

    finally:
        if driver:
            try:
                driver.quit()
            except OSError:
                pass
        print("✅ Navegador encerrado")


if __name__ == "__main__":
    main()