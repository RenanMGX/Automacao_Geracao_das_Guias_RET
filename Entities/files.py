import os
import pandas as pd
from abc import ABC, abstractmethod
from tkinter import filedialog
import xlwings as xw
from xlwings.main import App,Book,Sheet, Range, RangeRows, RangeColumns
import re
from .dependencies.functions import Functions
import asyncio
from typing import List, Dict

class NotFoundSheetError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
class ListOfDataEmptyError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
class FilePathEmpty(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
class PeriodoApuracaoNotFound(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)

class FilesManipulate():
    @property
    def file_path(self) -> str:
        return self.__file_path
    
    @property
    def df(self) -> pd.DataFrame:
        try:
            return self.__df
        except AttributeError:
            return pd.DataFrame()
        
    @property
    def periodo_apuracao(self) -> str:
        return self.__periodo_apuracao
    
    
        
    async def read_excel(self, file_path:str="", *, onlyValid: bool=False, visible:bool=False):
        if not file_path:
            file_path = filedialog.askopenfilename()
        
        if file_path == "":
            raise FilePathEmpty("o caminho do arquivo está vazio!")
            
                
        if os.path.exists(file_path):
            self.__file_path:str = file_path
        else:
            raise FileNotFoundError(f"Arquivo não encontrado '{file_path}'")
        
        if (self.file_path.endswith("xlsx")) or (self.file_path.endswith("xls")):
            await Functions.fechar_excel(self.file_path)
            if not onlyValid:
                await self.__extract_data(visible=visible)
        else:
            raise TypeError(f"Apenas arquivos em excel")
        
        return self
        
    async def __extract_data(self, *, visible:bool=False):
        from xlwings._xlwindows import COMRetryObjectWrapper
        def identificar_sheet(wb:Book, *, pattern:str=r'APURAÇÃO RET - [0-9]{6}') -> Sheet|None:
            for sheet_name in wb.sheet_names:
                name = re.match(pattern, sheet_name)
                if name:
                    return wb.sheets[wb.sheet_names.index(name.group())]
            return None       
        
        self.app:App = xw.App(visible=visible)
        self.wb:Book = self.app.books.open(self.file_path)
        if (result:=identificar_sheet(self.wb)):
            self.ws:Sheet = result
        else:
            raise NotFoundSheetError("a Sheet da 'APURAÇÃO' não foi encontrada!")
        
        #### Initial config #####
        self.initial_line:int = 14
        self.initial_column:str = "A"
        self.final_line:int = 2000
        self.final_column:str = "AG"
        #########################
        
        
        if (value:=re.search(r'[0-9]{6}', self.ws.name)):
            self.__periodo_apuracao:str = value.group()
        else:
            raise PeriodoApuracaoNotFound(f"não foi possivel identifica o periodo de apuração pelo nome da sheet '{self.ws.name}'")
        
        
        range_cells:Range = self.ws.range(f"{self.initial_column}{self.initial_line}:{self.final_column}{self.final_line}")
        
        
        cell_negatives:list = []
        count_none:int = 0
        columns_alimentar:Dict[str,list] = {"GIA4":[], "GIA1":[]}
        for row in range_cells.rows:
            row:Range
            if count_none > 10:
                break
            if all([x.value is None for x in row]):
                count_none += 1
                columns_alimentar['GIA4'].append(" ")
                columns_alimentar['GIA1'].append(" ")
            else:
                count_none = 0
                columns_alimentar['GIA4'].append(f'AF{row.row+1}')
                columns_alimentar['GIA1'].append(f'AG{row.row+1}')
                
                for cell in row.columns:
                    cell:Range
                    if cell.value:
                        if cell.api.Interior.Color == 11323383.0:
                            try:
                                if cell.value > 0:
                                    cell.value = -float(cell.value)
                                    cell_negatives.append(cell.address)
                            except:
                                pass
        
        
        self.ws.range(f"AF{self.initial_line}").value = "RPA_report - Guia 4%"
        self.ws.range(f"AF{self.initial_line+1}").value = [[addr] for addr in columns_alimentar["GIA4"]]
                 
        self.ws.range(f"AG{self.initial_line}").value = "RPA_report - Guia 1%"           
        self.ws.range(f"AG{self.initial_line+1}").value = [[addr] for addr in columns_alimentar["GIA1"]]
        
        #import pdb; pdb.set_trace()
        range_cells.row
        list_data:list|None = range_cells.value
        for cell_addres in cell_negatives:
            cell = self.ws.range(cell_addres)
            if cell.value:
                cell.value = -cell.value
        
        #import pdb; pdb.set_trace()  
        self.ws.range(f'AF15:AF{self.final_line}').value = ""
        self.ws.range(f'AG15:AG{self.final_line}').value = ""
              
        if list_data:
            df:pd.DataFrame = pd.DataFrame(list_data)
            df.columns = df.iloc[0]
            df = df.loc[:,~df.columns.duplicated()]
            df = df.drop(0)
            df = df.drop_duplicates()
            df = df[[
                                "Empresa",
                                "CNPJ RET",
                                "Valor a recolher 4%",
                                "Valor a recolher 1%",
                                "RPA_report - Guia 4%",
                                "RPA_report - Guia 1%" 
                            ]]
            df_4pc = df.drop(["Valor a recolher 1%", "RPA_report - Guia 1%"], axis=1)
            df_4pc = df_4pc.rename(columns={"Valor a recolher 4%":"Valor", "RPA_report - Guia 4%": "RPA_report"})
            df_4pc["Tipo"] = "Valor a recolher 4%"
            
            df_1pc = df.drop(["Valor a recolher 4%",  "RPA_report - Guia 4%"], axis=1)
            df_1pc = df_1pc.rename(columns={"Valor a recolher 1%":"Valor", "RPA_report - Guia 1%": "RPA_report"})
            df_1pc["Tipo"] = "Valor a recolher 1%"
            
            df = pd.concat([df_1pc, df_4pc], ignore_index=True)
            df = df[~df['Empresa'].isnull()]
            
            df = df[df['Valor'] > 0]
            
            for linha, value in df.iterrows():
                valor = round(value['Valor'], 2)
                valor = str(valor).replace('.', ',')
                df.loc[linha, 'Valor'] = valor #type: ignore
            
            #import pdb; pdb.set_trace()  
            self.__df:pd.DataFrame = df
            return self
        else:
            raise ListOfDataEmptyError("a lista de dados está vazia")
        
    async def record_return(self, *, value:str, address:str) -> None:
        if (_address:=re.search(r'[A-z]+[0-9]+', address)):
            try: 
                self.ws
                self.ws.range(_address.group()).value = str(value)
            except AttributeError:
                raise Exception(f"o arquivo precisa ser iniciado executando o metodo '{self.__class__.__name__}.read_excel()'")
        else:
            raise Exception(f"{address=} não é valido")
        
    async def close_excel(self, *, save:bool=False):
        try:
            if save:
                self.wb.save()
            self.wb.close()
            self.app.kill()
            
            await Functions.fechar_excel(self.file_path)
            
            del self.wb
            del self.app
        except Exception as error:
            print(error)
            
if __name__ == "__main__":
    pass       
