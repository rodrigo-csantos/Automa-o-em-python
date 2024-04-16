from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException
import time
import pyautogui
from features import utilities
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def accessWebsite(driver):
    driver.maximize_window()
    driver.get('URL do site aqui')

def login(driver):
    campo_usuario = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="user"]')))
    campo_usuario.send_keys('Usuário')
    campo_senha = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]')))
    campo_senha.send_keys('Senha')
    botao_entrar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="acessar"]')))
    botao_entrar.click()

def openSearchPage(driver):
    driver.get('URL da página aqui')

def downloadNFE(driver, NFE, fileName):
    print(NFE)
    time.sleep(1.5)
    sendNfe = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="notaFiscal"]')))
    sendNfe.click()
    sendNfe.send_keys(NFE)
    sendNfe.send_keys(Keys.ENTER)
    time.sleep(3)

    openPage = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//td[@valign='top']/a[contains(@href, 'consulta_minuta.php?chave=')]")))
    openPage.click()

    openPdf = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//tr[@nf='{NFE}']"))) 
    openPdf.click()
    time.sleep(1.5)

    num_guia_antes = len(driver.window_handles)
    
    try:
        downloadPDF = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="baixarDanfe"]')))
    except NoSuchElementException:
        try:
            downloadPDF = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="print_dec"]')))
        except NoSuchElementException:
            print("Nenhum dos elementos foi encontrado na página.")

    downloadPDF.click()
    time.sleep(1.5)
    
    num_guia_depois = len(driver.window_handles)

    if num_guia_depois > num_guia_antes:
        pyautogui.click(692,134, duration= 1.5)
        pyautogui.click(98,183, duration= 1.5)
        pyautogui.click(425,375, duration= 1.5)
        pyautogui.write(str(NFE))
        pyautogui.click(550,442, duration= 1.5)
        pyautogui.click(781,16, duration= 1.5)
        time.sleep(1.5)
        utilities.moveFile(utilities.pasta_origem, utilities.pasta_destinoPDF, ".pdf", fileName)
        time.sleep(1.5)
        filePath = os.path.join(utilities.pasta_destinoPDF, fileName)
        ContainsRefrigerated = utilities.filterRefrigeratedPDF(filePath)
        return ContainsRefrigerated
                
    else:
        keyNFE = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="chave_nfe"]')))
        valueNFE = keyNFE.get_attribute("value")
        return valueNFE

def exportExcelUndelivered (driver, initialDate, finalDate):
    data_in = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="data_inicial"]')))
    data_fi = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="data_final"]')))

    data_in.click()
    data_in.send_keys(initialDate)
    time.sleep(0.5)
    data_fi.click()
    data_fi.send_keys(finalDate)
    time.sleep(0.5)

    dropdown_situacao = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="situacao"]')))
    opcoes = Select(dropdown_situacao)
    opcoes.select_by_visible_text('SEM RECEBEDOR')

    pesquisar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="PESQUISAR"]')))
    pesquisar.click()
    
    excel = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="lista-minuta-excel"]')))
    excel.click()

def exportExcelDelivered (driver, initialDate, finalDate):
    data_in = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="data_inicial"]')))
    data_fi = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="data_final"]')))

    data_in.click()
    data_in.send_keys(initialDate)
    time.sleep(0.5)
    data_fi.click()
    data_fi.send_keys(finalDate)
    time.sleep(0.5)

    dropdown_situacao = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="situacao"]')))
    opcoes = Select(dropdown_situacao)
    opcoes.select_by_visible_text('ENTREGUE')

    pesquisar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="PESQUISAR"]')))
    pesquisar.click()

    excel = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="lista-minuta-excel"]')))
    excel.click()