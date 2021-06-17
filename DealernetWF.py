from selenium import webdriver
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as CondicaoEsperada
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import *
import os
import time

class CursoAutomacao:
    def __init__(self):
        chrome_options = Options()
        chrome_options.binary_location = os.getcwd() + os.sep + 'chrome-win'+ os.sep + 'chrome.exe'
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_argument('--lang=PT-BR')
        self.driver = webdriver.Chrome(executable_path=r'./chromedriver.exe',options=chrome_options)
        self.wait = WebDriverWait(
            driver=self.driver,
            timeout=10,
            poll_frequency=1,
            ignored_exceptions=[NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
            TimeoutException]
        )
        self.driver.implicitly_wait(10)
        self.site = 'https://grupoadtsa.dealernetworkflow.com.br/LoginAux.aspx?Windows'
    
    def Iniciar(self):
        self.driver.get(self.site)
        login_user = self.driver.find_element_by_xpath('//*[@id="vUSUARIO_IDENTIFICADORALTERNATIVO"]')
        login_pwd = self.driver.find_element_by_xpath('//*[@id="vUSUARIOSENHA_SENHA"]')
        login_btn = self.driver.find_element_by_xpath('//*[@id="IMAGE3"]')
        login_user.click()  
        login_user.send_keys('')
        login_pwd.click()
        login_pwd.send_keys('')
        login_btn.click()
        menu_selecionado = self.wait.until(CondicaoEsperada.element_to_be_clickable((By.XPATH,'//td[@class="x-toolbar-left"]//table//tbody//tr//td[8]')))
        #inserir_class = self.wait.until(CondicaoEsperada((By.XPATH,'//table[@id="TABLEACTIONS"]//tbody//tr//td')))
        frame1 = self.driver.find_element_by_xpath('//iframe[@src="../wwcalculoabc.aspx"]')
        self.driver.switch_to_frame(frame1)
        inserir_class = self.driver.find_element_by_xpath('//*[@id="INSERT"]')
        inserir_class.click()
        Agrupamentos = self.driver.find_element_by_xpath('//select[@id="vAGRUPAMENTO_CODIGO"]')
        OpcAgrupamentos = Select(Agrupamentos)
        OpcAgrupamentos.select_by_value("2")
        ClaABC = self.wait.until(CondicaoEsperada.visibility_of_element_located((By.XPATH,'//select[@id="vCALCULOABC_CLASABCCOD"]')))
        OpcClaABC = Select(ClaABC)
        OpcClaABC.select_by_value("1")



curso = CursoAutomacao()
curso.Iniciar()