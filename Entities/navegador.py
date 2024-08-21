from selenium import webdriver as navegador
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
#from time import sleep
from .dependencies.functions import P
import os
from getpass import getuser
from typing import List
import asyncio
import re
from time import sleep
from .dependencies.logs import Logs
import traceback

class WebElementNotFoundError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
class NavStartError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
class RegistersEmpty(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
class EmpresaNotFound(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
class ContribuinteError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)


def _find_element(by:str, target:str, *, driver:WebDriver|WebElement, timeout:int=5, force:bool=False) -> WebElement:
        for _ in range(timeout*4):
            try:
                return driver.find_element(by, target)
            except:
                asyncio.create_task(asyncio.sleep(.25))
        
        if force:
            return driver.find_element(By.TAG_NAME, 'html')
        
        raise WebElementNotFoundError(f"o elemento '{by}: {target}' não foi encontrado") 

def _find_elements(by:str, target:str, *, driver:WebDriver|WebElement, timeout:int=5, force:bool=False) -> list:
    for _ in range(timeout*4):
        try:
            return driver.find_elements(by, target)
        except:
            asyncio.create_task(asyncio.sleep(.25))
            
    if force:
        return []
    
    raise WebElementNotFoundError(f"o elemento '{by}: {target}' não foi encontrado") 

class SicalcReceita:
    @property
    def nav(self) -> WebDriver:
        try:
            return self.__nav
        except AttributeError:
            raise Exception(f"o navegador precisa ser iniciado executando o metodo '{self.__class__.__name__}.start()'")
    @nav.deleter
    def nav(self) -> None:
        try:
            del self.__nav
        except:
            pass
    
    async def start(self, *, initial_page:str="https://sicalc.receita.economia.gov.br/sicalc/rapido/contribuinte", restart_page:bool=False) -> bool:
        try:
            self.__nav
            if restart_page:
                self.nav.get(initial_page)
                return True
            print(P("O navegador já esta aberto!", color='red'))
            return False
        except AttributeError:
            print(P("Abrindo Navegador", color='blue'))
            self.__nav:WebDriver = await self.__start_nav(initial_page)
            print(P("O navegador aberto!", color='green'))
            return True
    
    async def __start_nav(self, url:str, *, download_path:str="downloads", timeout:int=10) -> WebDriver:
        if download_path:
            if not os.path.exists(download_path):
                download_path = os.path.join(os.getcwd(), download_path)
                os.makedirs(download_path)
            
            download_path = os.path.join(os.getcwd(), download_path)
            for file in os.listdir(download_path):
                file = os.path.join(download_path, file)
                if os.path.isfile(file):
                    try:
                        os.unlink(file)
                    except:
                        await Logs().register(status='Error',description='erro ao apagar arquivo', exception=traceback.format_exc())
            
            prefs:dict = {"download.default_directory": download_path}
            chrome_options: Options = Options()
            chrome_options.add_experimental_option('prefs', prefs)
            #chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            #chrome_options.add_experimental_option('useAutomationExtension', False)
            #chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_argument(f"--user-data-dir=C:\\Users\\{getuser()}\\AppData\\Local\\Google\\Chrome")
            #chrome_options.add_argument('--profile-directory=Renan Oliveira')
            #print(chrome_options.)
            #import pdb; pdb.set_trace()
            
            
        for _ in range(timeout):
            nav = navegador.Chrome(options=chrome_options)
            try:
                nav.get(url)
                return nav
            except:
                nav.close()
                await asyncio.sleep(1)
        raise NavStartError("não foi possivel iniciar o navegador")
    
    async def verificar_cadastros(self) -> List[str]:
        
        #_find_element(By.ID, "optionPJ", driver=self.nav).click()
        select:WebElement = _find_element(By.ID, 'selectToken', driver=self.nav)
        options:List[WebElement] = select.find_elements(By.TAG_NAME, 'option')
        
        if len(options) > 1:
            return [option.text for option in options]
        else:
            raise RegistersEmpty("Não foi encontrado empresas registradas")
    
    async def fechar(self) -> None:
        self.nav.close()
        del self.nav
        
    async def gerar_guia(self , *, cnpj:str, periodo_apuracao:str, valor:str):

        if not (valid:=re.search(r'[0-9]{2}.[0-9]{3}.[0-9]{3}/[0-9]{4}-[0-9]{2}', cnpj)):
            raise TypeError(f"numero de CNPJ invalido '{cnpj}'")
        else:
            if cnpj != valid.group():
                raise TypeError(f"numero de CNPJ invalido '{cnpj}'")
        
        select:WebElement = _find_element(By.ID, 'selectToken', driver=self.nav)
        options:List[WebElement] = select.find_elements(By.TAG_NAME, 'option')
        
        def find_empresa(cnpj):
            for option in options:
                if cnpj in option.text:
                    option.click()
                    return
            raise EmpresaNotFound(f"Não foi possivel encontrar a empresa com o CNPJ '{cnpj}'")
        find_empresa(cnpj=cnpj)
        
        botoes:WebElement = _find_element(By.ID, 'divBotoes', driver=self.nav)
        inputs:List[WebElement] = _find_elements(By.TAG_NAME, 'input', driver=botoes)
        for input in inputs:
            if input.get_attribute('value') == 'Continuar':
                input.click()
                break
   
        def erro_contribuinte() -> bool:
            try:
                _find_element(By.ID, 'contribuinte.errors', driver=self.nav)
                return True
            except:
                return False
        if erro_contribuinte():
            error_text:str = _find_element(By.ID, 'contribuinte.errors', driver=self.nav).text
            raise ContribuinteError(error_text)

        #_find_element(By.ID, 'observacao', driver=self.nav).clear()
        _find_element(By.ID, 'observacao', driver=self.nav).send_keys("RET")
        
        def find_autocomplete(target:str):
            auto_complete_s:List[WebElement] = _find_elements(By.CLASS_NAME, 'autocomplete-suggestion', driver=self.nav)            
            if not auto_complete_s:
                raise Exception("auto complete está vazio")
            for auto_complete in auto_complete_s:
                try:
                    if target in str(auto_complete.get_attribute('data-val')):
                        auto_complete.click()
                        return
                except:
                    continue
            raise Exception("auto complete não encontrado")

        def select_autocomplete():
            for _ in range(10):
                try:
                    _find_element(By.ID, 'codReceitaPrincipal', driver=self.nav).clear()
                    _find_element(By.ID, 'codReceitaPrincipal', driver=self.nav).send_keys('4095')
                    sleep(.5)
                    find_autocomplete('4095 - 01')
                    return
                except:
                    sleep(.5)
            raise Exception("auto complete não encontrado")
        
        select_autocomplete()
        sleep(.5)
        
        fld_automatico:WebElement = _find_element(By.ID, 'fldAutomatico', driver=self.nav)
        fld_automatico.find_element(By.ID, 'dataPA').clear()
        fld_automatico.find_element(By.ID, 'dataPA').send_keys(periodo_apuracao)
        
        fld_automatico.find_element(By.ID, 'numeroReferencia').click()
        
        if (periodo_error:=_find_element(By.ID, 'fldError', driver=self.nav).text) != '':
            raise Exception(periodo_error)
        
        
        fld_principal:WebElement = _find_element(By.ID, 'fldPrincipal', driver=self.nav)
        
        while len(str(fld_principal.find_element(By.ID, 'valorPrincipal').get_attribute('value'))) > 0:
            fld_principal.find_element(By.ID, 'valorPrincipal').send_keys(Keys.BACKSPACE)
        
        
        fld_principal.find_element(By.ID, 'valorPrincipal').send_keys(valor)
        
        _find_element(By.ID, 'btnCalcular', driver=self.nav).click()
        
        tbody = _find_element(By.TAG_NAME, 'tbody', driver=self.nav)
        sleep(1)
        tbody.find_element(By.TAG_NAME, 'input').click()
        
        _find_element(By.ID, 'btnDarf', driver=self.nav).click()
        
        self.nav.window_handles
        
        import pdb; pdb.set_trace()
        
        return
    
if __name__ == "__main__":
    pass
