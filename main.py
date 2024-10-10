from Entities.files import FilesManipulate, pd
from Entities.navegador import SicalcReceita, TimeoutException, NoSuchElementException, InvalidSessionIdException, JavascriptException
from typing import List, Dict, Literal, Coroutine, Any
from Entities.dependencies.functions import P
from Entities.interface import Ui_Interface, QtWidgets
from Entities.dependencies.logs import Logs
import sys
import qasync
import asyncio
import traceback
from datetime import datetime


class Execute(Ui_Interface):
    @property
    def log(self) -> Logs:
        return Logs(name="Automação Geração das Guias RET")
    
    @property
    def excel_file(self) -> FilesManipulate:
        return self.__excel_file
    @excel_file.setter
    def excel_file(self, value:FilesManipulate) -> None:
        if isinstance(value, FilesManipulate):
           self.__excel_file:FilesManipulate = value
           return
        raise TypeError("Tipo do Arquivo Invalido")
    @excel_file.deleter
    def excel_file(self) -> None:
        try:
            del self.__excel_file
        except:
            pass
    
    @property
    def navegador(self) -> SicalcReceita:
        return self.__navegador
    
    @property
    def file_manipulate(self) -> FilesManipulate:
        return self.__file_manipulate

    def __init__(self) -> None:
        #################################################################################################################################
        super().__init__(version="1.6") # <------------------------------------------------ Sempre alterar versão antes de compilar     #
        #################################################################################################################################
        self.__file_manipulate:FilesManipulate = FilesManipulate()
        self.__navegador: SicalcReceita = SicalcReceita()
    
    def closeEvent(self, event):
        print(P("Encerrando Script"))
        try:
            self.navegador.nav.close()
            del self.navegador.nav
        except:
            pass
        try:
            asyncio.create_task(self.file_manipulate.close_excel())
        except:
            pass
        event.accept()  # Isso irá fechar a janela    
        
        
    def pg02_action_voltar(self) -> None:
        async def voltar_async(self:Execute):
            await self.pg01_print_aviso(reset=True)
            await self.pg02_print_infor(reset=True)
            await self.pg02_list_limpar_items()
            await self.mudar_pagina('Inicial')
            del self.excel_file
            
        asyncio.create_task(voltar_async(self))
        
    async def initial_config(self):
        self.version = "1.0"
        self.pg01_bt_carregar_arquivo.clicked.connect(self.carregar_excel)
        self.pg02_bt_verific_empre.clicked.connect(self.fazer_verificacao_empresas)
        self.pg02_bt_voltar.clicked.connect(self.pg02_action_voltar)
        self.pg02_bt_iniciar.clicked.connect(self.iniciar_gerar_guias)
        #self.telas.setCurrentIndex(1)
        
    def teste(self):
        asyncio.create_task(self.pg02_list_limpar_items())
    
    def carregar_excel(self):  
        async def carregar_excel_async(self:Execute) -> None :
            await self.pg01_print_aviso(reset=True)
            await self.pg01_bt_carregar_arquivo_visibilidade(False)
            try:
                self.excel_file = await self.__file_manipulate.read_excel(onlyValid=True)
                await self.mudar_pagina('Pos-Inicial')
            except Exception as error:
                await self.pg01_print_aviso(text=str(error), color='red')
            finally:
                await self.pg01_bt_carregar_arquivo_visibilidade(True)
                return

        asyncio.create_task(carregar_excel_async(self))
    
    def fazer_verificacao_empresas(self):
        self.empresas_verificadas_sem_cadastro:list = []
        async def start_async(self:Execute):
            mensagem_final = ""
            asyncio.create_task(self.pg02_bt_verific_visibilidade(False))
            asyncio.create_task(self.pg02_bt_iniciar_visibilidade(False))
            await self.pg02_print_infor(reset=True)
            try:
                await self.pg02_list_limpar_items()
                await self.pg02_print_infor(text="Iniciando Verificação")
                
                self.excel_file
                self.excel_file = await self.__file_manipulate.read_excel(self.excel_file.file_path)
                #excel_file:FilesManipulate = self.__file_manipulate.read_excel()
                await self.pg02_print_infor(text="Abrindo Navegador")
                await self.navegador.start(restart_page=True)
                
                await self.pg02_print_infor(text="Listando Empresas")
                #self.navegador.start()
                empresas_verificadas = await self.verificar_empresas(self.excel_file.df)
                
                await self.pg02_print_infor(text="Listando Empresas")
                self.empresas_verificadas_sem_cadastro = list(set(empresas_verificadas['Sem Cadastro']))
                if self.empresas_verificadas_sem_cadastro:
                    for cnpj in self.empresas_verificadas_sem_cadastro:
                        await self.pg02_list_additem(str(cnpj))
                    mensagem_final += "existe empresas que não estão cadastradas no site\n"
                else:
                    mensagem_final += "nenhuma empresa pendente, pronto para iniciar\n"
                
                await self.pg02_print_infor(text="Encerrando verificação")
                
                mensagem_final += "Verificação Encerrada \n"
                await self.pg02_print_infor(text=mensagem_final)
                await self.pg02_bt_iniciar_visibilidade(True)
            except AttributeError as error_attribute:
                print(P(str(error_attribute), color='red'))
                print(traceback.format_exc())
                await self.pg02_print_infor(text=f"Planilha do Excel não foi carregada", color='red')
            except Exception as error:
                await self.pg02_print_infor(text=str(error))
            finally:
                await self.__excel_file.close_excel()
                await self.pg02_bt_verific_visibilidade(True)
                return
       
        asyncio.create_task(start_async(self))
                    
    async def verificar_empresas(self, df:pd.DataFrame) -> Dict[Literal["Com Cadastro", "Sem Cadastro"],List[str]]:
        result:Dict[Literal["Com Cadastro", "Sem Cadastro"],List[str]] = {"Com Cadastro": [], "Sem Cadastro": []}
        lista_empresas_cadastradas:List[str] = await self.__navegador.verificar_cadastros()
        for row, value in df.iterrows():
            cnpj:str = value['CNPJ RET']
            if cnpj:
                cnpj = cnpj.replace(" ", "")
                if cnpj in " - ".join(lista_empresas_cadastradas):
                    result["Com Cadastro"].append(cnpj)
                    #print(P(f"A empresa '{cnpj}' está cadastrada!", color='green'))
                else:
                    result["Sem Cadastro"].append(cnpj)
                    #print(P(f"A empresa '{cnpj}' não está cadastrada!", color='red'))
        return result
    
    def iniciar_gerar_guias(self):
        async def async_start(self:Execute):
            tempo_inicio = datetime.now()
            await self.pg02_print_infor(reset=True)
            await self.navegador.limpar_pasta_download()
            await self.pg02_bt_verific_visibilidade(False)
            await self.pg02_bt_iniciar_visibilidade(False)
                #await self.file_manipulate.read_excel(self.excel_file.file_path)
            for row, value in self.file_manipulate.df.iterrows():
                erro:Exception
                for _ in range(60):
                    if value["CNPJ RET"] in self.empresas_verificadas_sem_cadastro:
                        continue
                    try:
                        await self.navegador.start(speak=False)
                        await self.navegador.gerar_guia(cnpj=value["CNPJ RET"], periodo_apuracao=self.file_manipulate.periodo_apuracao, valor=value["Valor"])
                        #await self.file_manipulate.record_return(address=value["RPA_report"], value="Concluido")
                        await FilesManipulate.renomear_arquivo_recente(download_path=self.navegador.download_path, empresa=value['Empresa'], divisao=value['Divisão'], valor=value['Valor'], tipo=value['Tipo'])
                        print(P(f"{value["CNPJ RET"]} - foi concluido", color="green"))
                        break
                    except Exception as err:
                        asyncio.create_task(self.log.register(status='Report', description=str(err), exception=traceback.format_exc()))
                        try:
                            self.navegador.nav.close()
                        except:
                            pass
                        try:
                            del self.navegador.nav
                        except:
                            pass
                        await self.navegador.start()
                        await asyncio.sleep(1)                               
                        # except TimeoutException as error:
                        #     erro = error
                        #     await asyncio.sleep(1)
                        # except NoSuchElementException as error:
                        #     erro = error
                        #     await asyncio.sleep(1)
                        # except InvalidSessionIdException as error:
                        #     erro = error
                        #     del self.navegador.nav
                        #     await self.navegador.start()
                        #     await asyncio.sleep(1)   
                        # except JavascriptException as error:
                        #     erro = error
                        #     await asyncio.sleep(1)                        
                        # except Exception as error:
                        #     asyncio.create_task(self.log.register(status='Report', description="erro ao gerar arquivo", exception=traceback.format_exc()))
                        #     await asyncio.sleep(1)
                        #     erro = error
                        #     if "Nenhuma conexão pôde ser feita porque a máquina de destino as recusou ativamente" in str(error):
                        #         print(P("Dechando Navegador e abrindo novamente"))
                        #         del self.navegador.nav
                        #         await self.navegador.start()
                        #         await asyncio.sleep(1)
                        #     #await self.file_manipulate.record_return(address=value["RPA_report"], value=error)
                        #     #break
                        print(P(f"Erro com '{value["CNPJ RET"]}' tentando novamente -> {type(err)}", color="red"))

            await self.pg02_bt_verific_visibilidade(True)
            #await self.pg02_bt_iniciar_visibilidade(True)
            await self.file_manipulate.close_excel(save=True)
            print(P("Fim da Automação!", color='green'))
            await self.pg02_print_infor(text=f"Fim da Automação!\ntempo de execução: {datetime.now() - tempo_inicio}")
            print(P(f"tempo de execução: {datetime.now() - tempo_inicio}", color='white'))
            asyncio.create_task(self.log.register(status='Concluido', description="execução bem sucedida"))
            self.navegador.nav.close()
            del self.navegador.nav
        asyncio.create_task(async_start(self))
                    
    @staticmethod
    async def iterador(df: pd.DataFrame) -> dict:
        result:dict = {}
        for row, value in df.iterrows():
            result[row] = value
        return result
    
         
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)
    #MainWindow = QtWidgets.QMainWindow()
    ui = Execute()
    asyncio.run(ui.setupUi())
    asyncio.run(ui.initial_config())
    ui.show()
    loop.run_forever()
