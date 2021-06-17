from selenium import webdriver
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as CondicaoEsperada
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import *
from datetime import datetime
from pathlib import Path
import os
import time

class CursoAutomacao:
    def __init__(self):
        chrome_options = Options()
        chrome_options.binary_location = os.getcwd() + os.sep + 'chrome-win'+ os.sep + 'chrome.exe'
       #chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'], )
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory":  "C:\ETL",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })
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
        menu_selecionado = self.wait.until(CondicaoEsperada.element_to_be_clickable((By.XPATH,'/html/body/div[1]/table/tbody/tr/td[1]/table/tbody/tr/td[8]')))
        menu_selecionado.click()
        time.sleep(3)
        menu_relatorio = self.driver.find_element_by_xpath('//li[contains(@id, "Relatórios")]')
        ActionChains(self.driver).key_down(Keys.CONTROL).move_to_element(menu_relatorio).perform()
        time.sleep(2)
        menu_AcompanhamentoProdutivos = self.driver.find_element_by_xpath('//li[contains(@id, "AcompanhamentodeProdutivo")]')
        ActionChains(self.driver).key_down(Keys.CONTROL).move_to_element(menu_AcompanhamentoProdutivos).perform()
        menu_vendasperiodo = self.driver.find_element_by_xpath('//li[contains(@id, "FaturamentodeServiçosporProdutivo")]')
        ActionChains(self.driver).key_down(Keys.CONTROL).click(menu_vendasperiodo).perform()
        time.sleep(2)
        frame1 = self.driver.find_element_by_xpath('//iframe[@src="../wp_relfaturamentoservicosprodutivos.aspx"]')
        self.driver.switch_to_frame(frame1)
        data = datetime.now().strftime("01%m%Y")
        dataInicial = self.driver.find_element_by_xpath('//*[@id="vDATAINICIO"]')
        dataInicial.send_keys(Keys.DELETE)
        dataInicial.send_keys(data)
        dataInicial.send_keys(Keys.TAB)
        Agrupamentos = self.driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/span/input[2]')
        Agrupamentos.click()
        gerarRelatorio = self.driver.find_element_by_xpath('//*[@name="BTNGERAR"]')
        gerarRelatorio.click()
        time.sleep(5)
        self.driver.switch_to.default_content()
        frameMl = self.driver.find_element_by_xpath('//iframe[@src="../wp_reportingservices.aspx"]')
        #self.driver.switch_to_frame(frameMl)
        if frameMl is not None:
            print("Encontrou algo")
            self.driver.switch_to_frame(frameMl)
        arquivos = Path('C:\ETL\REL_FaturamentoServicosProdutivos_Report.xlsx')
        if arquivos is not None:
            try:
                arquivos.unlink()
            except OSError as e:
                print(f"Error:{ e.strerror}")
        ExportarMD = self.wait.until(CondicaoEsperada.element_to_be_clickable((By.XPATH,'//*[@name="IMGMALADIRETA"]')))
        ExportarMD.click()
        
        os.system('C:\Pentaho\pdi-ce-9.0.0.0-423\data-integration\Kitchen.bat /file:C:\ETL\RVD.kjb')
        self.driver.close()
       
        
        
       
        #menu_relatorio.click()
        #inserir_class = self.wait.until(CondicaoEsperada((By.XPATH,'//table[@id="TABLEACTIONS"]//tbody//tr//td')))
    



curso = CursoAutomacao()
curso.Iniciar()