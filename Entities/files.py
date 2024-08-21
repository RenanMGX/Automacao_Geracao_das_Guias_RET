import os
import pandas as pd
from abc import ABC, abstractmethod
from tkinter import filedialog
import xlwings as xw
from xlwings.main import App,Book,Sheet, Range, RangeRows, RangeColumns
import re
from .dependencies.functions import Functions
import asyncio
from typing import List

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
        
        self.__app:App = xw.App(visible=visible)
        self.__wb:Book = self.__app.books.open(self.file_path)
        if (result:=identificar_sheet(self.__wb)):
            self.__ws:Sheet = result
        else:
            raise NotFoundSheetError("a Sheet da 'APURAÇÃO' não foi encontrada!")
        
        #### Initial config #####
        initial_line:str = "14"
        initial_column:str = "A"
        final_line:str = "2000"
        final_column:str = "AG"
        #########################
        
        #######################
        ws = self.__ws
        ######################
        
        if (value:=re.search(r'[0-9]{6}', self.__ws.name)):
            self.__periodo_apuracao:str = value.group()
        else:
            raise PeriodoApuracaoNotFound(f"não foi possivel identifica o periodo de apuração pelo nome da sheet '{self.__ws.name}'")
        
        
        range_cells:Range = self.__ws.range(f"{initial_column}{initial_line}:{final_column}{final_line}")
        
        
        cell_negatives:list = []
        count_none:int = 0
        # for cell in range_cells:
        #     if not cell.value:
        #         count_none += 1
        #     else:
        #         count_none = 0
        #     if count_none >= 1000:
        #         break
        #     if cell.api.Interior.Color == 11323383.0:
        #         if cell.value:
        #             try:
        #                 if cell.value > 0:
        #                     cell.value = -float(cell.value) 
        #                     cell_negatives.append(cell.address)
        #             except:
        #                 pass
        for row in range_cells.rows:
            row:Range
            if count_none > 10:
                break
            if all([x.value is None for x in row]):
                count_none += 1
            else:
                count_none = 0
                self.__ws.range(f"AF{row.row}").value = f'AF{row.row}'
                self.__ws.range(f"AG{row.row}").value = f'AG{row.row}'
                
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
                        
                                
        self.__ws.range(f"AF{initial_line}").value = "RPA_report - Guia 4%"           
        self.__ws.range(f"AG{initial_line}").value = "RPA_report - Guia 1%"           
                
        
        #import pdb; pdb.set_trace()
        range_cells.row
        list_data:list|None = range_cells.value
        for cell_addres in cell_negatives:
            cell = self.__ws.range(cell_addres)
            if cell.value:
                cell.value = -cell.value
                
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
            
            self.__df:pd.DataFrame = df
            return self
        else:
            raise ListOfDataEmptyError("a lista de dados está vazia")
        
    async def close_excel(self):
        try:
            self.__wb.close()
            self.__app.kill()
            
            await Functions.fechar_excel(self.file_path)
            
            del self.__wb
            del self.__app
        except Exception as error:
            print(error)
            
if __name__ == "__main__":
    pass       
