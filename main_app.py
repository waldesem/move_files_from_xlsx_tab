import shutil
import sqlite3
import sys
import os
from datetime import date, datetime

import openpyxl


# поля в таблице candidates базы данных candidates:
SQL_CAND = 'staff, department, full_name, last_name, birthday, birth_place, country, series_passport, ' \
           'number_passport, date_given, snils, inn, reg_address, live_address, phone, email, education, ' \
           'check_work_place, check_passport, check_debt, check_bankruptcy, check_bki, check_affiliation, ' \
           'check_internet, check_cronos, check_cross, resume, date_check, officer'
# поля в таблице registry базы данных candidates:
SQL_REG = 'fio, birthday, staff, checks, recruiter, date_in, officer, date_out, result, final_date, url'
# CONNECT = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\candidates.db' #     файл базы данных
# MAIN_FILE = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'  # файл кандидатов
# WORK_DIR = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\\'    # рабочая папка кандидатов
# DESTINATION = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Персонал\Персонал-2\\'   # архивная папка
# INFO_FILE = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\Запросы по работникам.xlsx'  # запросы
# REPORT_FILE = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\Отчетность ЦКБ ДЭБ БКБ.xlsx'   # отчет

# тестовые константы:
CONNECT = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\candidates.db'  #     файл базы данных
MAIN_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'   # файл кандидатов
WORK_DIR = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\\'  # рабочая папка
DESTINATION = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Персонал\Персонал-2\\' # архивная папка
INFO_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Запросы по работникам.xlsx' # запросы
REPORT_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Отчетность ЦКБ ДЭБ БКБ.xlsx' # отчет


def check_modify(): # проверка главного файлa на изменения с датой сегодня
    if date.fromtimestamp(os.path.getmtime(MAIN_FILE)) != date.today():
        print('Исходный файл не изменялся')
        print('Работа завершена')
        sys.exit()

def backup(filename = (MAIN_FILE, CONNECT, INFO_FILE, REPORT_FILE)) -> None:    #  резервные копии файлов
    for file in filename:
        shutil.copy(file, DESTINATION)
    print('Созданы резервные копии')

def range_row(sheet) -> list: # получаем список номеров строк соответствующих сегодняшней дате
    row_num = []
    for cell in sheet['K5000':'K20000']:
        for c in cell:
            if isinstance (c.value, datetime):
                if (c.value).strftime('%Y-%m-%d') == date.today().strftime('%Y-%m-%d'):
                    row_num.append(c.row)
            elif str(c.value).strip() == date.today().strftime('%d.%m.%Y'):
                row_num.append(c.row)
    print('Создан список номеров строк')
    if len(row_num) == 0:
        print('Данные за сегодня не найдены')
        print('Работа завершена')
        sys.exit()
    return row_num

def dir_range(sheet, row_num) -> list:  # получаем список папок в директории рабочей директории
    fio = [sheet['B' + str(i)].value.strip().lower() for i in row_num]
    subdir = [sub for sub in os.listdir(WORK_DIR) if sub.lower().strip() in fio]
    if len(subdir) == 0: # если папок соответствующих записям в файле нет - передаем записи из реестра в БД
        print('Папки соответствующие записям не найдены')
        registry_check(row_num, sheet)
        print('Работа завершена')
        sys.exit()
    print('Создан список субдиректорий')
    return subdir

def file_range(book, sheet, subdir, row_num) -> list: # получаем список путей к файлам Заключений
    name_path = list(filter(None, []))
    for f in subdir:
        subdir_path = f"{WORK_DIR[0:-1] + f}\\"
        for file_name in os.listdir(subdir_path):
            if file_name.startswith("Заключение") and file_name.endswith("xlsm"):
                name_path.append(os.path.join(WORK_DIR, subdir_path, file_name))
    if len(name_path) == 0: # если файлов соответствующих шаблону нет, перемещаем папки, передаем данные в БД
        print('Файлы заключений не найдены')
        create_link(sheet, subdir, row_num)
        registry_check(row_num, sheet)
        save_workbook(book)
        print('Работа завершена')
        sys.exit()
    print('Получен список путей файлов заключений')
    return name_path

def check_types(check) -> list: # проверка и преобразование типов данных
    view = []
    for c in check:
        if isinstance(c, datetime):  # преобразование datetime в строку
            view.append(c.strftime('%Y-%m-%d'))
        else:
            view.append(str(c))    # преобразование в строку
    return view

def parse_excel(path_files) -> list: # получаем список кортежей с данными заключений
    values_zakl = []
    for path in path_files: # открываем Заключения с анкетой для чтения данных
        if len(path_files) != 0:
            wbz = openpyxl.load_workbook(path, keep_vba=True)
            wsz = wbz.worksheets[0]     # анкета из заключения:
            # staff, department, full_name, last_name, birthday, birth_place, country, series_passport, 
            # number_passport, date_given, snils, inn, reg_address, live_address, phone, email, education
            form_view = check_types([wsz['C4'].value, wsz['C5'].value, wsz['C6'].value, wsz['C7'].value, 
                                    wsz['C8'].value, 'None','None', wsz['C9'].value, wsz['D9'].value, 
                                    wsz['E9'].value, 'None', wsz['C10'].value, 'None', 'None', 'None', 
                                    'None', 'None'])
            if len(wbz.sheetnames) > 1:
                wsz = wbz.worksheets[1]
                if str(wsz['K1'].value) == 'ФИО':     # если есть анкета, перезаписываем переменную form_view
                    form_view = check_types([wsz['C3'].value, wsz['D3'].value, wsz['K3'].value, wsz['S3'].value, 
                                            wsz['L3'].value, wsz['M3'].value, wsz['T3'].value, wsz['P3'].value, 
                                            wsz['Q3'].value, wsz['R3'].value, wsz['U3'].value, wsz['V3'].value, 
                                            wsz['N3'].value, wsz['O3'].value, wsz['Y3'].value, wsz['Z3'].value, 
                                            wsz['X3'].value])
            wsz = wbz.worksheets[0] # Данные проверки: check_work_place, check_passport, check_debt, check_bankruptcy, 
            # check_bki, check_affiliation, check_internet, check_cronos, check_cross, resume, date_check, officer
            form_check = check_types([f"{wsz['C11'].value} - {wsz['D11'].value}; {wsz['C12'].value} - "
                                    f"{wsz['D12'].value}; {wsz['C13'].value} - {wsz['D13'].value}", 
                                    wsz['C17'].value, wsz['C18'].value, wsz['C19'].value, wsz['C20'].value,
                                    wsz['C21'].value, wsz['C22'].value, 
                                    f"{wsz['B14'].value}: {wsz['C14'].value}; {wsz['B15'].value}: {wsz['C15'].value}",
                                    wsz['C16'].value, wsz['C23'].value, wsz['C24'].value, wsz['C25'].value])
            values_zakl.append(tuple(form_view + form_check))   # формируем данные для запроса
            wbz.close()   # Закрываем книгу Excel
    print('Получены данные анкет и заключений')
    return values_zakl

def create_link(sheet, subdir, row_num) -> None:    # создаем гиперссылки и перемещаем папки
    for nums in row_num:
        for sub in subdir:
            if str(sheet['B' + str(nums)].value.strip().lower()) == sub.strip().lower():
                sbd = sheet['B' + str(nums)].value.strip()
                lnk = f"{DESTINATION + sbd[0][0]}\\{sbd} - {sheet['A' + str(nums)].value}"
                sheet['L' + str(nums)].hyperlink = lnk    # записывает в книгу
                shutil.move(WORK_DIR + sbd, lnk)
    print('Созданы гиперссылки. Файлы  успешно перенесены')

def insert_db(database, query, value) -> None:  # запись в БД
    with sqlite3.connect(database, timeout=5.0) as con:
        cur = con.cursor()
        if len(value) > 0:
            cur.executemany(query, value)
            print('Передан запрос в БД')
        else:
            print('Получен пустой запрос')

def registry_check(row_num, sheet) -> None:  # Получаем данные из реестра и передаем в БД
    reg_val = []
    for n in row_num: # получаем список значений строк соответствующих сегодняшней дате
        reg_val.append(tuple(check_types([c.value for cell in sheet['B' + str(n):'L' + str(n)] for c in cell])))
    ins_reg = "INSERT INTO registry ({}) VALUES ({}?)".format(SQL_REG, '?, ' * (len(SQL_REG.split()) - 1))
    print(reg_val)
    insert_db(CONNECT, ins_reg, reg_val)
    print('Данные перенесены в реестр')

def save_workbook(book):   # Сохраняем книгу Excel
    book.save(MAIN_FILE)
    print('Книга сохранена')

def main(): # главная функция
    wb = openpyxl.load_workbook(MAIN_FILE, keep_vba=True, read_only=False)
    ws = wb.worksheets[0]   # открываем первый лист книги MAIN_FILE для чтения и записи данных
    num_row = range_row(ws) # записываем номера строк соответствующих сегодняшней дате
    subdirectory = dir_range(ws, num_row)   # список директорий, которые соответствуют фамилиям кандидатов
    ins_cand = "INSERT INTO candidates ({}) VALUES ({}?)".format(SQL_CAND, '?, ' * (len(SQL_CAND.split()) - 1))
    insert_db(CONNECT, ins_cand, parse_excel(file_range(wb, ws, subdirectory, num_row)))   # передаем данные в БД
    create_link(ws, subdirectory, num_row)  # Создание гиперссылок и перемещение папок
    registry_check(num_row, ws) # получаем данные из реестра и передаем в БД
    save_workbook(wb)  # Сохраняем книгу Excel


if __name__ == "__main__":
    check_modify()
    backup()
    main()
