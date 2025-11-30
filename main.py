import pyexcel as p
import json, random
import time

import os

from docx import Document
from docxtpl import DocxTemplate

class ParserTables():
    """Обработчик таблиц (чтение и подстановка со сдвигом)"""
    def __init__(self):
        pass
    # example table:
    #[['№', 'наименование', 'Кол-во новое', 'Кол-во предыдующее', 'Прочее'], 
    # ['1', '2', '3', '4', '5'], 
    # ['1', 'Red', '11', '123', ''], 
    # ['2', 'Green', '22', '1234', ''], 
    # ['3', 'Blue', '33', '23456', ''], 
    # ['№', 'наименование', 'Кол-во новое', 'Кол-во предыдующее', 'Прочее'], 
    # ['1', '2', '3', '4', '5'], 
    # ['1', 'One', '11', '123', ''], 
    # ['2', 'Two', '22', '1234', ''], 
    # ['3', 'Three', '33', '23456', ''], 
    # ['4', 'Four', '44', '3456', ''], 
    # ['№', 'наименование', 'Кол-во новое', 'Кол-во предыдующее', 'Прочее'], 
    # ['1', '2', '3', '4', '5'], 
    # ['1', 'Bob', '11', '123', ''], 
    # ['2', 'Tom', '22', '1234', ''], 
    # ['3', 'Ada', '33', '23456', '']]
    """перезапись таблицы с использованием подстановки из файла сводной таблицы"""
    def RewriteAllTable(self, array_tables, table_entity_counters):
        table_entity_counters = table_entity_counters[0] # [[[ ]]] --> [[ ]]
        keys = {}
        for elem in table_entity_counters:
            keys[elem[0]] = elem[1]
        for i in range(len(array_tables)):
            array_tables[i][3] = array_tables[i][2] 
            if array_tables[i][1] in keys:
                array_tables[i][2] = keys[array_tables[i][1]]
            else:
                array_tables[i][2] = "not found"
        print(array_tables)
        return array_tables
        


class ScriptODS():
    """работа с файлом в формате .ods"""
    def __init__(self, file_name_sheet):
        self.name_sheet_entity_counter_out = "Sheet 0"
        self.file_name_sheet = file_name_sheet

    """создание тестовой таблицы со множеством страниц и данными"""
    def CreateFakeTable(self, arr_keys, count_elements_in_page, count_pages, file_name):     
        # (["red","green"], 5, 160000) 
        content = {}
        for i in range (count_elements_in_page):
            data = []
            for _ in range (count_pages):
                temp = [arr_keys[random.randint(0,len(arr_keys)-1)]]
                data.append(temp)
            content["Sheet "+str(i)]=data
        p.save_book_as(bookdict=content, dest_file_name=file_name)
            
    """получение из массива -> таблицы ключей и количества появлений"""
    def GetDataMapCountersElements(self, datamap):
        collect_datamap = [] # 3d-array [ [['Marta'],['Tom']], [['Marta'],['Bob']] ]

        for page in datamap:
            collect_datamap += page

        result_dict = {}
        for elem in collect_datamap:
            if len(elem) == 0:
                continue
            elem = elem[0]
            if elem not in result_dict:
                result_dict[elem] = 0
            result_dict[elem] += 1

        out = []
        for d in result_dict:
            out.append([d,result_dict[d]])
        return out  

    """сохранить словарь упоминаний в файл"""
    def SaveSheetToFile(self, datamap, file_name):
        book = {}
        book[self.name_sheet_entity_counter_out] = datamap
        p.save_book_as(bookdict = book, dest_file_name = file_name)
       
    """объединить все листы из одного файла в один лист и сохранить новым файлом"""
    def MergedSheetsFromFiles(self, array_files_names, file_name_out):
        datamap = [] # 2d-array [['Marta'], ['Tom'], ['Marta'], ['Bob'], ['Bob'], ['Bob']]
        for file_name in array_files_names:
            book = p.get_book(file_name = file_name)
            names = book.sheet_names()
            for name in names:  
                rows = book[name].get_array()  
                filtre_rows = rows #self.FilterFromIndexColumn(rows,0)  
                datamap += [filtre_rows]
            print("merge books:", self.GetDataMapCountersElements(datamap))
        #        return self.GetDataMapCountersElements(datamap)
        self.SaveSheetToFile(self.GetDataMapCountersElements(datamap), file_name_out)

    """фильтр колонок - оставить только колонки в строках с индексом N (где N >= 0)"""
    def FilterFromIndexColumn(self,array,column):
        filtre_array = []
        for row in array:
            filtre_array += row[column]
        return filtre_array
    
    """получить 3d массив из таблицы"""
    def GetTableEntityCounters(self):
        book = p.get_book(file_name = self.file_name_sheet)
        names = book.sheet_names()
        datamap = []
        for name in names:  
            rows = book[name].get_array()  
            filtre_rows = rows #self.FilterFromIndexColumn(rows,0)  
            datamap += [filtre_rows]
        print("merge books:", self.GetDataMapCountersElements(datamap))
        return datamap

        

class ScriptDOCX():
    """работа с файлом в формате .docx"""
    def __init__(self):
        self.context = {
            'var_1':'999', 
            'var_2':'qwerty'
        }    

    """получить 3d-массив все строк из всех таблиц из файла .docx"""
    def ExtractTextFromAllTables(self, file_name):
        temp_document = Document(file_name)
        collect_rows_from_tables = []
        for temp_table in temp_document.tables:
            for temp_row in temp_table.rows:
                collect_rows_from_tables.append([cell.text.strip() for cell in temp_row.cells])
        return collect_rows_from_tables

    """подстановка данных из словаря в шаблон и сохранение"""
    def RenderTemplateFromContext(self, file_name, context):
        doc = DocxTemplate("input.docx")
        doc.render(context)
        doc.save("output1.docx")


class ScriptConvertFileExt():
    """конвертация файла между форматами .docx <--> .odt"""
    def __init__(self):
        self.cmd_odt_to_docx = "libreoffice --headless --convert-to docx"
        self.cmd_docx_to_odt = "libreoffice --headless --convert-to odt"

    """конвертировать .docx --> .odt"""
    def DocxToOdt(self, file_name_input_docx):
        os.system(f"{self.cmd_docx_to_odt} {file_name_input_docx}")        

    """конвертировать .odt --> .docx"""
    def OdtToDocx(self, file_name_input_odt):
        os.system(f"{self.cmd_odt_to_docx} {file_name_input_odt}")


class ExecutionLogic():
    """управление логикой выполнения скриптов"""
    def __init__(self):
        self.array_files_names_ods = [
            "file_1.ods",
            "file_2.ods",
            "file_3.ods",        
        ]
        self.file_template_docx = "file_template.docx"
        self.file_preview_week_odt = "file_preview_week.odt"
        self.file_preview_week_docx = "file_preview_week.docx"
        self.file_new_week_odt = "file_new_week.odt"
        self.file_name_out_all_sheets = "file_out.ods"
    
        self.script_ods = ScriptODS(self.file_name_out_all_sheets)
        self.script_docx = ScriptDOCX()
        self.script_convert_file_ext = ScriptConvertFileExt()
        self.parser_tables = ParserTables()
 
    """точка входа"""
    def Main(self):
        self.script_ods.CreateFakeTable(["Red","Green","Blue"],10,3, "file_1.ods")
        self.script_ods.CreateFakeTable(["One","Two","Three","Four"],10,3, "file_2.ods")
        self.script_ods.CreateFakeTable(["Bob","Tom","Ada"],10,3, "file_3.ods")

        self.CreateFileEntityCounter(self.array_files_names_ods, self.file_name_out_all_sheets)
        self.GetPreviewValueFromTable(self.file_preview_week_odt, self.file_preview_week_docx)
        
    
    """создание файла с одной сводной таблицей из перечисления файлов"""
    def CreateFileEntityCounter(self,array_files_names, file_name_out):
        self.script_ods.MergedSheetsFromFiles(array_files_names, file_name_out)
    
    """Извлечь предыдующие значения из всех таблиц (подготовка к перенесению в другую колонку)"""
    def GetPreviewValueFromTable(self, file_preview_week_odt, file_preview_week_docx):
        self.script_convert_file_ext.OdtToDocx(file_preview_week_odt)
        res = self.script_docx.ExtractTextFromAllTables(file_preview_week_docx)
        new_tables = self.parser_tables.RewriteAllTable(res, self.script_ods.GetTableEntityCounters())
        
        


exec_logic = ExecutionLogic()
exec_logic.Main()







