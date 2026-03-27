import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def clicar_hcaptcha(driver, timeout=10, use_coordinates=False, x=None, y=None):
    """
    Clica no checkbox do hCaptcha (ou reCAPTCHA) na página atual.
    
    Parâmetros:
        driver: instância do WebDriver.
        timeout: tempo máximo de espera (segundos).
        use_coordinates: se True, usa pyautogui para clicar em coordenadas absolutas.
        x, y: coordenadas (obrigatório se use_coordinates=True).
        
    Retorna:
        True se o clique foi realizado, False caso contrário.
    """
    if use_coordinates:
        # Abordagem por coordenadas (simula clique humano)
        try:
            import pyautogui
            pyautogui.click(x, y)
            print(f"✅ hCaptcha clicado via coordenadas ({x}, {y})")
            return True
        except ImportError:
            print("❌ pyautogui não instalado. Instale com: pip install pyautogui")
            return False
        except Exception as e:
            print(f"❌ Erro ao clicar por coordenadas: {e}")
            return False

    # Abordagem via iframe (mais robusta)
    try:
        # Aguarda o iframe do hCaptcha
        iframe = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[contains(@title, 'hCaptcha') or contains(@title, 'recaptcha')]"))
        )
        driver.switch_to.frame(iframe)
        print("✅ Mudou para o iframe do captcha")

        # Clica no checkbox usando JavaScript (funciona mesmo com shadow DOM)
        driver.execute_script("document.querySelector('[role=\"checkbox\"]').click();")
        print("✅ Checkbox clicado via JavaScript")

        driver.switch_to.default_content()
        return True

    except Exception as e:
        print(f"❌ Falha ao clicar no captcha via iframe: {e}")
        driver.switch_to.default_content()  # Garante que volta ao conteúdo principal
        return False