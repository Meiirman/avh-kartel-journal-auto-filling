'''
# OTDScript_V1 -Script для ОТДЕЛА ТЕХНИЧЕСКОЙ ДОКУМЕНТАЦИИ версия 1
# СКРИПТ ДЛЯ АВТОЗАПОЛНЕНИЯ ДОКУМЕНТОВ ДЛЯ ОТДЕЛА ТЕХНИЧЕСКОЙ ДОКУМЕНТАЦИИ

Цели:
    -Автоматизировать однотипные задачи отдела.
    -Сократить время обработки документов

Бизнес Требования к скрипту:
    - Скрипт должен читать INPUT данные
    - Записывать их на нужные места в документах
    - Сохраниять документ с записанными данными

    
План выполнения работ:

1. Нстроить шаблоны
    - шблон "1 Журнал авторского надзора"
    - шблон "2 Журнал сварочных работ"
    - шблон "3 Акт скрытых работ"
    - шблон "4 Журнал работ по МСК"
    - шблон "5 Журнал производства работ"
2. Написать скрипт который прочитает данные из Excel
3. Написать скрипт заполняющий файлы данными:
    - Скрипт для "1 Журнал авторского надзора"
    - Скрипт для "2 Журнал сварочных работ"
    - Скрипт для "3 Акт скрытых работ"
    - Скрипт для "4 Журнал работ по МСК"
    - Скрипт для "5 Журнал производства работ"
4. Тестирование
5. Стилизация
6. Конвертирование в exe файл
7. Тест в формате EXE
8. Обучение пользователей
9. Доработка
10. Отчет

'''

import traceback
import openpyxl
import datetime
from docxtpl import DocxTemplate
from pyautogui import alert
import os

import requests

TEMPLATES_FOLDER_PATH = "Шаблоны/"
OFP = ""
INPUT_FILE_PATH = "INPUT.xlsx"
INPUT_FILE_PAGE_NAME = "Основные переменные"

TEMPLATES = [
    { 
        "TEMPLATE" : TEMPLATES_FOLDER_PATH+"1 Журнал авторского надзора.docx",
        "FILE_NAME" :     "Журнал авторского надзора"
    },{ 
        "TEMPLATE" : TEMPLATES_FOLDER_PATH+"2 Журнал сварочных работ.docx",
        "FILE_NAME" :     "Журнал сварочных работ"
    },
    { 
        "TEMPLATE" : TEMPLATES_FOLDER_PATH+"Акт скрытых работ/3 Акт скрытых работ.docx",
        "FILE_NAME" :     "Акт скрытых работ"
    },
    { 
        "TEMPLATE" : TEMPLATES_FOLDER_PATH+"4 Журнал работ по МСК.docx",
        "FILE_NAME" : "Журнал работ по МСК"
    },{ 
        "TEMPLATE" : TEMPLATES_FOLDER_PATH+"5 Журнал производства работ.docx",
        "FILE_NAME" : "Журнал производства работ"
    }
]

OUTPUT_FILE_NAMES = [
    "Журнал авторского надзора",
    "Журнал сварочных работ",
    "Акт скрытых работ",
    "Журнал работ по МСК",
    "Журнал производства работ"
]


def none_to_empty_string(dictionary):
    months = {
         1 : "январь",
         2 : "феврать",
         3 : "март",
         4 : "апрель",
         5 : "май",
         6 : "июнь",
         7 : "июль",
         8 : "август",
         9 : "сентябрь",
        10 : "октябрь",
        11 : "ноябрь",
        12 : "декабрь"
        
    }
    data_dict = {key: value if value is not None else "" for key, value in dictionary.items()}
    for key, value in data_dict.items():
        print(f'{key} = {value}')
        try:
            data_dict[key] = f'«{value.day if value.day >= 10 else f"0{value.day}"}» {months[value.month]} {value.year}г.'       
            if "date_and_num_" in str(key) or "tmpl2_tbl1_sdate" in str(key) or  "tmpl2_tbl1_edate" in str(key) or  "tmpl2_tbl3_warkdate" in str(key) or  "asr_one_sdate_q" in str(key) or  "asr_one_edate_q" in str(key) or  "tmpl45_tbl1_sdate_q" in str(key) or  "tmpl45_tbl1_edate_q" in str(key) :  
                
                data_dict[key] = f'{value.day if value.day >= 10 else f"0{value.day}"}.{value.month if value.month >= 10 else f"0{value.month}"}.{value.year}'       
                # print(f'key:{key}, val:{data_dict[key]}')``

            # print(data_dict[key])
        except:
            # print("ds")
            pass

    return data_dict


def get_excel_data():
    try:
    # if True:
        # Открыть файл и взять данные из нужного листа 
        sheet = openpyxl.load_workbook(INPUT_FILE_PATH, data_only=True)[INPUT_FILE_PAGE_NAME]
        
       # Присваивание в data_dict значение из листа 
        data_dict = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if len(row) >= 2:
                key, value = row[:2]
                data_dict[key] = value


        # Удаление лишних строк
        for i in range(1,31):
            try:
                if  data_dict[f'subcontractor_{i}'] == None: 
                    data_dict[f'date_and_num_{i}'] = None
                    data_dict[f'did_not_have_{i}'] = None
                    data_dict[f'tmpl1_representativ_{i}'] = None
                    data_dict[f'default_value_{i}'] = None
                    data_dict[f'company_name_{i}'] = None
            except:
                pass     


        for i in range(1,31):
            try:
                if data_dict[f'ASR__akt_num_{i}'] == None:
                    data_dict[f'aktosrnumb_{i}'] = None
                print(f"aktosrnumb_{i} = {data_dict[f'ASR__akt_num_{i}']}")
            except:
                pass

        for i in range(1,31):
            try:
                if data_dict[f'tmpl45_tbl1_fpos_{i}'] == None:
                    data_dict[f'tmpl45_tbl1_sdate_q{i}'] = None
                    data_dict[f'tmpl45_tbl1_edate_q{i}'] = None
                print(f"tmpl45_tbl1_sdate_q{i} = {data_dict[f'tmpl45_tbl1_edate_q{i}']}")
            except:
                pass
        

        # Превратить все None в пустой string ""
        data_dict = none_to_empty_string(data_dict)
        
        # Удалить название словаря None
        del data_dict[None]
        del data_dict[0]

        return data_dict
    # Обработчик ошибок
    except Exception as e:
        alert(f"Ошибка в Ексель файле\nError message: {e} (EXCEL)")
        return None


def replace_words_in_docx(dictionary, file_path, file_name, OUTPUT_FOLDER_PATH=OFP):
    try:
    # if True:
        output_file_path = f"{OUTPUT_FOLDER_PATH}{file_name} {dictionary['obj_name']}.docx"

        # Берет шаблон, заполняет данными из словаря, сохраняет как новый файл
        print(file_path)
        doc = DocxTemplate(file_path)
        doc.render(dictionary)
        doc.save(output_file_path)

    except Exception as e:
        print(f"Ошибка в шаблоне Word файла\nError message: {e} (WORD)")
        alert(f"Ошибка в шаблоне Word файла\nError message: {e} (WORD)")
        return None





def preparatory_operations():
    pass


import os
from docx import Document



def merge_docx_files(input_files, output_file):
    merged_doc = Document()

    for file in input_files:
        doc = Document(file)

        for element in doc.element.body:
            merged_doc.element.body.append(element)

    merged_doc.save(output_file)
    print(f"Файлы успешно объединены и сохранены в '{output_file}'.")




def data_assignments():

    data = get_excel_data()
    OUTPUT_FOLDER_PATH = data['obj_name'] + "/"
        
    try:
        os.mkdir(OUTPUT_FOLDER_PATH)
    except:
        pass
    for tmpl in TEMPLATES:
        # alert(tmpl['FILE_NAME'])
        if tmpl['FILE_NAME'] == "Акт скрытых работ":
            count = 0 
            for i in range(1,30):
                try:
                    param_1 = data['ASR__akt_name_'+str(i)]
                    param_2 = data['ASR__akt_num_'+str(i)]
                    param_3 = data['ASR__akt_deskrop_'+str(i)]
                    param_4 = data['ASR__akt_useditems_'+str(i)]

                    param_5 = data['asr_one_sdate_'+str(i)]
                    param_6 = data['asr_one_edate_'+str(i)]
                    param_7 = data['ASR__akt_name_'+str(i+1)]

                    if param_7 == "":
                        param_7 = "Установка ФБС блоков и ЦМ контейнера"
                        count = i+1


                    XXX = {"param_1":param_1,"param_2":param_2,"param_3":param_3,"param_4":param_4,"param_5":param_5,"param_6":param_6,"param_7":param_7}                   
                    print(XXX)
                    data.update(XXX)

                    doc = DocxTemplate(TEMPLATES_FOLDER_PATH+"Акт скрытых работ/3 Акт скрытых работ.docx")
                    doc.render(data)
                    doc.save(TEMPLATES_FOLDER_PATH+f"Акт скрытых работ/file{i}.docx")
                    if param_7 == "Установка ФБС блоков и ЦМ контейнера":
                        break
                except Exception as erroroerrer:
                    print(erroroerrer)
            
            file_names = []
            for i in range(1,count):
                file_names.append(TEMPLATES_FOLDER_PATH+f"Акт скрытых работ/file{i}.docx")
            merge_docx_files(file_names, f"{OUTPUT_FOLDER_PATH}{tmpl['FILE_NAME']} {data['obj_name']}.docx")
            
            file_names = []
            for i in range(1,30):
                file_names.append(TEMPLATES_FOLDER_PATH+f"Акт скрытых работ/file{i}.docx")
            for file_name in file_names:
                try:
                    os.remove(file_name)
                    print(f"Файл '{file_name}' успешно удален.")
                except FileNotFoundError:
                    print(f"Файл '{file_name}' не найден.")
                except Exception as e:
                    print(f"Произошла ошибка при удалении файла '{file_name}': {e}")
                
        else:
            replace_words_in_docx(data, tmpl['TEMPLATE'], tmpl['FILE_NAME'], OUTPUT_FOLDER_PATH)

def send_report(text=None, process=None, responsible=None):
    requests.post(f"https://script.google.com/macros/s/AKfycbzDwjE6Pu1a7otho2EHwbI-4yNoEmLijTfwWfI3toWpDpJ6rc-O1pKljV6XMLJmQIyJ/exec?time={datetime.datetime.now().strftime('%d.%m.%Y %H:%M:%S')}&process={process}&responsible={responsible}&text={text}")


if __name__ == "__main__":
    try:
        preparatory_operations()
        data_assignments()

        send_report(text="Автозаполнение журналов(ИД) для Кар-тел", process="Автозаполнения журналов(ИД)", responsible=os.getlogin())
        alert("Конец процесса")
    except:
        alert(f"Произошла ошибка: {traceback.format_exc()}")



