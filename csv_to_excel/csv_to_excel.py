#!/usr/bin/python3
# -*- coding: utf-8 -*-

import csv
import logging
import os

from xlsxwriter.workbook import Workbook  # Для записи в .xlsx

fh_log = [logging.FileHandler("csv_to_excel.log", 'w', 'utf-8')]
logging.basicConfig(handlers=fh_log, format="%(message)s", level=logging.INFO)


def converter(csvfiles_list):
    for csvfile in csvfiles_list:
        logging.info(f"Конвертирование {csvfile}")
        workbook = Workbook(csvfile.replace("csv", "xlsx"))  # Создание файла .xlsx с таким же именем
        worksheet = workbook.add_worksheet()  # Создание страницы в таблице
        try:
            with open(csvfile, encoding='utf-8') as f:  # Открытие на чтение
                reader = csv.reader(f, delimiter=';', quotechar='|')
                for r, row in enumerate(reader):
                    for c, col in enumerate(row):
                        worksheet.write(r, c, col)
            workbook.close()
            logging.info(f"Файл {csvfile} конвертирован.\n")
        except (UnicodeDecodeError, PermissionError) as e:
            logging.info(f"Ошибка конвертирование:\n\t{e}\n")


env_edunetroot = os.getenv("ISI_isinquirer_root", "/home/chda/DEVELOP/edunet")  # Путь до edunet
files = filter(lambda x: x[-4:] == '.csv', os.listdir(env_edunetroot))  # Все файлы .csv, которые найдем
csvfiles_list = [os.path.join(env_edunetroot, x) for x in files]  # Список путей к файлам .csv
if csvfiles_list:
    logging.info("Найдены следующие .csv файлы:")
    for csvfile in csvfiles_list:
        logging.info(f"\t{csvfile}")
    logging.info("\n")
    converter(csvfiles_list)
else:
    logging.info(".csv файлы не найдены.\n")
logging.info("Выход.\n")
