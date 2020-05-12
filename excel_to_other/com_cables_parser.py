# -*- coding: utf-8 -*-

import logging
import xlrd # Импорт библиотеки для работы с .xlsx (excel)
from re import fullmatch

# Функция обработки повтора в ячейках
def repeat_check(last_indiv_text, current_text):
    if current_text == "": # Отдельное условие для пустых ячеек 
        return ""
    elif last_indiv_text == current_text: # Условие для совпадения
        return "-//-"
    else:
        return current_text


# def valid_check(row_list):
#     print(row_list)
#     print(fullmatch(r"^\d+$", row_list[0]))
#     print(fullmatch(r"^[^/&{}#]+$", row_list[1]))
#     print(fullmatch(r"^\d+\.\d+-\d+/\d+:(0\d|[^0]\d+)$", row_list[2]))
#     print(fullmatch(r"^([A-Z]+\.)?\d+\.\d+$|^резерв$", row_list[3]))
#     print(fullmatch(r"^\d+$", row_list[5]))
#     print("\n")

# Функция поиска и икранирования специальных символов Latex, если они есть в списке-строчки
def esc_spec_symbols(row_list):
    output_row_list = row_list.copy()
    spec_symbols_list = ["\\", "&", "{", "}", "#", "%"] # Список спец. символов, расширяемый
    for i in range(len(output_row_list)):
        for spec_symbol in spec_symbols_list:
            ss_index = output_row_list[i].find(spec_symbol) # Вывод индекса спец. символа, поиск начинается с начала строки
            while ss_index != -1: # -1: спец. символы не найдены
                output_row_list[i] = output_row_list[i][0:ss_index] + "\\" + output_row_list[i][ss_index:]
                ss_index = output_row_list[i].find(spec_symbol, ss_index+2) # Поиск следующего спец. символа, поиск начинается после прошлого найденного (+2: \%)
    return output_row_list

# Функция поиска нужного листа в .xlsx
def search_desired_sheet(table): 
    for sheet_name in table.sheet_names(): # Перебор по всем именам листов
        if fullmatch("^(.+[ _-])?([Oo]ut|[Pp]arser)([ _-].+)?$", sheet_name) != None: # Примернаый RE
            sheet = table.sheet_by_name(sheet_name) # После прохождению по RE имени, проверяем данные в листе по шапке
            row_list = sheet.row_values(0)
            if (
                    row_list[0] == "№" and
                    fullmatch(r"^[Нн]азначение[ _-]+порта$", row_list[1]) and
                    fullmatch(r"^[Нн]омер[ _-]+порта[ _-]+(СКС|скс)$", row_list[2]) and
                    fullmatch(r"^[Рр]озетки$", row_list[3]) and
                    fullmatch(r"^[Сс]етевое[ _-]+оборудование$", row_list[4]) and
                    fullmatch(r"^[Нн]омер[ _-]+порта[ _-]+коммутатора$", row_list[5]) and
                    fullmatch(r"^[Мм]арка[ _-]+кабеля$", row_list[6])
                ):
                print('Найден лист с данными таблицы коммутации.')
                return sheet # Cсылка на страницу, не данные
            else:
                print('Лист "' + sheet_name + '" не имеет данных таблицы коммутации. Пропускаем.')
        else:
            print('Лист "' + sheet_name + '" не соответствует по маске имени. Пропускаем.')
    print('Ошибка. Лист с данными таблицы коммутации не найден.')
    return None # Как и fullmatch, буду возвращать None, если все названия не соответствовали RE или везде не те данные

# Функция преобразования данных float в str, если такие есть
def convers_table_row(row_list):
    for x in (0, 5):
        if row_list[x]:
            row_list[x] = '%.f' % row_list[x]
    return row_list

logging.basicConfig(filename="sample.log", level=logging.INFO)


# table = xlrd.open_workbook("./test.xlsx") # Открытие таблицы на чтение
table = xlrd.open_workbook("./Таблица коммутации.xlsx") # Открытие таблицы на чтение
sheet = search_desired_sheet(table)
if sheet:
    last_indiv_str = [""] * 7 # Место хранения значений прошлой строки
    # Создание файла с данными для tex
    # Если файл уже создан, он будет перезаписан
    with open('com-cabel-lines.tex.inc', 'w') as f:
        for row_num in range(1, sheet.nrows):
            row = [repeat_check(i, j) for i, j in zip(last_indiv_str, convers_table_row(sheet.row_values(row_num)))]
            row = esc_spec_symbols(row) # Генератор строки для записи в файл с учетом условий повторения
            f.write(r"\lengformat{%s} &\codeformat{%s} & \nameformat{%s} & \noteformat{%s} & \codeformat{%s} & \codeformat{%s} & \codeformat{%s} \\\hline" % tuple(row) + "\n")
            last_indiv_str = sheet.row_values(row_num)
