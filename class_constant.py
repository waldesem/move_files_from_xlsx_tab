import sqlite3
import openpyxl
from datetime import date


# сегодняшняя дата
DATE_TODAY = date.today().strftime('%Y-%m-%d') + ' 00:00:00'
# поля в таблице candidates базы данных personal
SQL_CAND = 'staff, department, full_name, last_name,	birthday, birth_place, country, series_passport, ' \
           'number_passport, date_given, snils, inn, reg_address, live_address, phone, email, education, ' \
           'check_work_place, check_passport, check_debt, check_bankruptcy, check_bki, check_affiliation, ' \
           'check_internet, check_cronos, check_cross, check_rand_info, resume, date_check, officer'
# поля в таблице registry базы данных personal
SQL_REG = 'fio, birthday, staff, checks, recruiter, date_in, officer, date_out, result, final_date, url'
# файл базы данных где находится реестр и результаты проверки
CONNECT = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\candidates.db'
# главный файл кандидатов
MAIN_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'
# рабочая папка кандидатов
WORK_DIR = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\\'
# место хранения отработанных кандидатов
DESTINATION = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Персонал\Персонал-2\\'


class Database:
    """Объявляем класс Database для работы с базой данных"""

    def __init__(self, database, query, value):
        self.database = database
        self.query = query
        self.value = value

    # функция для записи в БД
    def insert_db(self):
        try:
            with sqlite3.connect(self.database) as con:
                cur = con.cursor()
                cur.execute(self.query, self.value)
                con.commit()
        except sqlite3.Error as error:
            print('Ошибка', error)


class ExelFile:
    """Объявляем класс для работы с таблицами Excel"""

    def __init__(self, exel_file):
        self.exel_file = exel_file

    # открытие листа с данными Excel
    def open_file(self, sheet):
        workbook = openpyxl.load_workbook(self.exel_file, keep_vba=True)
        worksheet = workbook.worksheets[sheet]
        return worksheet

    # сохранение листа с данными Excel
    def save(self, exel_file):
        self.exel_file = exel_file
        workbook = openpyxl.load_workbook(self.exel_file, keep_vba=True)
        workbook.save(exel_file)
