from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pyautogui

driver = webdriver.Chrome()

try:
    # Acessar site
    driver.get("https://sisparnet.pgfn.fazenda.gov.br/sisparInternet/internet/darf/consultaParcelamentoDarfInternet.xhtml")
    driver.maximize_window()
    time.sleep(5)
    
    # Digitar CNPJ
    driver.find_element(By.ID, "formConsultaParcelamentoDarf:imNrCpfCnpjParcelamento").send_keys("10480769000108")
    time.sleep(1)
    
    # TAB
    driver.switch_to.active_element.send_keys(Keys.TAB)
    time.sleep(1)
    
    # Digitar número
    driver.switch_to.active_element.send_keys("9059903")
    time.sleep(1)
    
    # Clicar no captcha
    pyautogui.click(232, 462)
    print("✅ Captcha clicado!")
    
    # Aguardar mais tempo para o captcha processar
    print("Aguardando 5 segundos para processamento do captcha...")
    time.sleep(5)
    
    # Verificar se o captcha foi marcado (opcional)
    try:
        checkbox = driver.find_element(By.XPATH, "//div[@role='checkbox']")
        is_checked = checkbox.get_attribute("aria-checked")
        print(f"Status do captcha: aria-checked={is_checked}")
    except:
        pass
    
    # Clicar no botão Consultar
    driver.find_element(By.ID, "formConsultaParcelamentoDarf:btnConsultar").click()
    print("✅ Botão Consultar clicado!")
    
    time.sleep(10)
    
finally:
    driver.quit()