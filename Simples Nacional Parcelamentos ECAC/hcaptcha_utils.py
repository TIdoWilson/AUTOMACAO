import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def resolver_hcaptcha(driver, timeout=20):
    """
    Resolve o checkbox do hCaptcha (invisível ou visível) na página atual.
    Retorna True se conseguiu clicar, False caso contrário.
    """
    try:
        # Aguarda o iframe do hCaptcha
        iframe = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'hcaptcha')]"))
        )
        driver.switch_to.frame(iframe)
        print("✅ Mudou para iframe do hCaptcha")

        # Tenta clicar no checkbox (pode estar visível ou invisível)
        try:
            checkbox = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@role='checkbox']"))
            )
            checkbox.click()
            print("✅ Checkbox do hCaptcha clicado")
        except:
            driver.execute_script("document.querySelector('[role=\"checkbox\"]').click();")
            print("✅ Checkbox invisível clicado via JS")

        driver.switch_to.default_content()
        time.sleep(1)  # aguarda o captcha resolver
        return True

    except Exception as e:
        print(f"❌ Falha ao resolver hCaptcha via iframe: {e}")
        driver.switch_to.default_content()
        return False