import shutil
import sqlite3
import os
from datetime import date, datetime

from conclude import *


DATE = date.today() # сегодняшняя дата (объект)
CONNECT = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\candidates.db' #     файл базы данных
WORK_DIR = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\\'    # рабочая папка кандидатов
DESTINATION = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Персонал\Персонал-2\\'   # архивная папка
INFO_FILE = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\Запросы по работникам.xlsx'  # запросы
MAIN_FILE = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'  # файл кандидатов
# поля в таблице candidates базы данных candidates:
SQL_CAND = 'staff, department, full_name, last_name, birthday, birth_place, country, series_passport, ' \
           'number_passport, date_given, snils, inn, reg_address, live_address, phone, email, education, ' \
           'check_work_place, check_passport, check_debt, check_bankruptcy, check_bki, check_affiliation, ' \
           'check_internet, check_cronos, check_cross, resume, date_check, officer'
# поля в таблице registry базы данных candidates:
SQL_REG = 'fio, birthday, staff, checks, recruiter, date_in, officer, date_out, result, final_date, url'
# поля в таблице inquiry базы данных candidates:
SQL_INQ = 'full_name, birthday, staff, period, info, firm, date_inq'
# тестовые константы:
# CONNECT = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\candidates.db'  #     файл базы данных
# WORK_DIR = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\\'  # рабочая папка
# DESTINATION = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Персонал\Персонал-2\\' # архивная папка
# MAIN_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'   # файл кандидатов
# INFO_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Запросы по работникам.xlsx' # запросы


def range_row(sheet): # get list of row nums that correspond with today date
    row_num = []
    for cell in sheet:
        for c in cell:
            if isinstance (c.value, datetime): # check format of date 
                if c.value.strftime('%Y-%m-%d') == DATE.strftime('%Y-%m-%d'):
                    row_num.append(c.row)
            elif str(c.value).strip() == DATE.strftime('%d.%m.%Y'):
                row_num.append(c.row)
    return row_num

def parse_conclusions(sheet, num_row):
    subdirectory = dir_range(sheet, num_row)  # list of directories with candidates names
    if len(subdirectory):
        path_files = file_range(subdirectory)  # create list with paths of conclusions
        if len(path_files):
            ins_cand = "INSERT INTO candidates ({}) VALUES ({}?)".format(SQL_CAND, '?, ' * (len(SQL_CAND.split()) - 1))
            insert_db(ins_cand, parse_excel(path_files))  # parse files and send info to database
        create_link(sheet, subdirectory, num_row)  # create url and move folders to archive

def dir_range(sheet, row_num):  # получаем список папок в рабочей директории
    fio = [sheet['B' + str(i)].value.strip().lower() for i in row_num]
    subdir = [sub for sub in os.listdir(WORK_DIR) if sub.lower().strip() in fio]
    return subdir

def file_range(subdir): # получаем список путей к файлам Заключений
    name_path = list(filter(None, []))
    for f in subdir:
        subdir_path = os.path.join(WORK_DIR[0:-1], f)
        for file in os.listdir(subdir_path):
            if file.startswith("Заключение") and (file.endswith("xlsm") or file.endswith("xlsx")):
                name_path.append(os.path.join(WORK_DIR, subdir_path, file))
    return name_path

def check_types(checks): # convert datatypes to string
    view = [c.strftime('%Y-%m-%d') if isinstance(c, datetime) else str(c).strip() for c in checks]
    return view

def parse_excel(path_files): # получаем список кортежей с данными заключений
    conclusion = []
    for path in path_files: # открываем Заключения с анкетой для чтения данных
        if len(path_files):
            form = Forms(path)  # формируем данные для запроса
            conclusion.append(tuple(check_types(form.get_conclusion())))   
    return conclusion

def create_link(sheet, subdir, row_num):    # создаем гиперссылки и перемещаем папки
    for n in row_num:
        for sub in subdir:
            if str(sheet['B' + str(n)].value.strip().lower()) == sub.strip().lower():
                sbd = sheet['B' + str(n)].value.strip()
                lnk = os.path.join(DESTINATION, sbd[0][0], f"{sbd} - {sheet['A' + str(n)].value}")
                sheet['L' + str(n)].hyperlink = str(lnk)    # записывает в книгу
                shutil.move(os.path.join(WORK_DIR, sbd), lnk)

def insert_db(query, value):  # запись в БД
    with sqlite3.connect(CONNECT, timeout=5.0) as con:
        cur = con.cursor()
        if len(value):
            cur.executemany(query, value)

def registry_check(row_num, sheet, start, end, table, columns) -> None:  # Получаем данные из реестра и передаем в БД
    reg_val = []
    for n in row_num: # получаем список значений строк соответствующих сегодняшней дате
        reg_val.append(tuple(check_types([c.value for cell in sheet[start + str(n):end + str(n)] for c in cell])))
    reg_query = "INSERT INTO {} ({}) VALUES ({}?)".format(table, columns, '?, ' * (len(columns.split()) - 1))
    insert_db(reg_query, reg_val)
