import shutil
import sqlite3
import sys
from datetime import date
import os

import openpyxl
from openpyxl import Workbook


# поля в таблице candidates базы данных candidates
SQL_CAND = 'staff, department, full_name, last_name, birthday, birth_place, country, series_passport, ' \
           'number_passport, date_given, snils, inn, reg_address, live_address, phone, email, education, ' \
           'check_work_place, check_passport, check_debt, check_bankruptcy, check_bki, check_affiliation, ' \
           'check_internet, check_cronos, check_cross, resume, date_check, officer'
# поля в таблице registry базы данных candidates
SQL_REG = 'fio, birthday, staff, checks, recruiter, date_in, officer, date_out, result, final_date, url'
# файл базы данных где находится реестр и результаты проверки
CONNECT = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\candidates.db'
# CONNECT = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\candidates.db'
# главный файл кандидатов
MAIN_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'
# MAIN_FILE = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'
# рабочая папка кандидатов
WORK_DIR = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\\'
# WORK_DIR = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\\'
# место хранения кандидатов
DESTINATION = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Персонал\Персонал-2\\'
# DESTINATION = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Персонал\Персонал-2\\'


class Database:
    """Объявляем класс Database для работы с базой данных"""

    def __init__(self, database, query, value):
        self.database = database
        self.query = query
        self.value = value

    # метод для записи в БД
    def insert_db(self):
        with sqlite3.connect(self.database, timeout=5.0) as con:
            cur = con.cursor()
            if type(self.value) is tuple:   # построчная запись в БД
                cur.execute(self.query, self.value)
            elif type(self.value) is list:  # групповая запись в БД
                cur.executemany(self.query, self.value)


class BackUp:
    """Объявляем класс BackUp для копирования данных"""

    def __init__(self, current, destination):
        self.current = current
        self.destination = destination

    # метод для копирования данных
    def backup(self):
        shutil.copy(self.current, self.destination)

    # метод для переноса данных
    def remove(self):
        shutil.move(self.current, self.destination)


class ExcelFile(Workbook):
    """Объявляем класс ExcelFile для работы с таблицами Excel, наследуем класс Workbook"""

    def __init__(self, excel_file):
        super().__init__(excel_file)
        self.num_row = None
        self.worksheet = None
        self.workbook = openpyxl.load_workbook(excel_file, keep_vba=True, read_only=False)

    # метод для открытия таблиц
    def open_sheet(self, sheet):
        self.worksheet = self.workbook.worksheets[sheet]
        return  self.worksheet

    # метод для закрытия файла
    def close_workbook(self):
        self.workbook.close()

    # метод для сохранения файла
    def save_workbook(self, excel_file):
        self.workbook.save(excel_file)

    # метод для разбора ячеек и поиска номера строк по текущей дате
    def range_row(self):
        self.num_row = []
        for cell in self.worksheet['K5000':'K20000']:
            for c in cell:
                if str(c.value) == date.today().strftime('%Y-%m-%d') + ' 00:00:00':
                    self.num_row.append(c.row)
        if len(self.num_row) == 0:  # если список пустой - выходим из программы
            sys.exit()
        return self.num_row

    # метод записи значений строк, соответствующих номеру строки
    def reg_range(self):
        reg_val = []
        for n in self.num_row:
            reg_val.append(tuple(map(str, [c.value for cell in ws['B' + str(n):'L' + str(n)] for c in cell])))
        return reg_val


class DirSubdir:
    """Объявляем класс FileDir для работы с файлами"""

    def __init__(self, directory):
        self.name = None
        self.subdirectory = None
        self.directory = directory

    # метод по перебору папок в директории
    def dir_range(self, name):
        self.name = name
        subdirectory = []
        for s in os.listdir(self.directory):
            if s.lower().strip() in self.name:
                subdirectory.append(s)
        if len(subdirectory) == 0:  # если список пустой - выходим из программы
            sys.exit()
        return subdirectory

    # метод по перебору файлов по шаблону
    def file_range(self, subdirectory):
        self.subdirectory = subdirectory
        name_path = ''
        for file_name in os.listdir(self.subdirectory):
            if file_name.startswith("Заключение") and file_name.endswith("xlsm"):
                name_path = os.path.join(WORK_DIR, self.subdirectory, file_name)
        return name_path


class HyperLink:
    """Объявляем класс HyperLink для создания и записи гиперссылок"""

    def __init__(self):
        self.num_rom = row_num
        self.subdir = subdir
        self.ws = ws

    # метод создает гиперссылки
    def create_link(self):
        h_link = {}
        for nums in self.num_rom:
            if str(self.ws['B' + str(nums)].value) in self.subdir:
                sbd = self.ws['B' + str(nums)].value.strip()
                lnk = f"{DESTINATION + sbd[0][0]}\\{sbd} - {self.ws['A' + str(nums)].value}"
                self.ws['L' + str(nums)].hyperlink = lnk    # записывает в книгу
                h_link[sbd] = lnk   # создает словарь: подпапка = путь
        return h_link

class Forms:
    """Объявляем класс Forms для работы с данными"""
    
    def __init__(self):
        self.form_view = None
        self.form_check = None
        self.cand_values = None

    # анкетные данные
    def view_form(self, staff, department, full_name, last_name, birthday, birth_place, country, series_passport,
                  number_passport, date_given, snils, inn, reg_address, live_address, phone, email, education):

        self.form_view = {'staff':staff, 'department': department, 'full_name': full_name,
                          'last_name': last_name, 'birthday':birthday, 'birth_place': birth_place,
                          'country': country, 'series_passport': series_passport, 'number_passport': number_passport,
                          'date_given': date_given, 'snils': snils, 'inn': inn,
                          'reg_address': reg_address, 'live_address': live_address, 'phone': phone,
                          'email': email, 'education': education
                        }
        return self.form_view

    # данные заключения
    def check_form(self,check_work_place, check_cronos, check_cross, check_passport, check_debt, check_bankruptcy,
                   check_bki, check_affiliation, check_internet, resume, date_check, officer):
        self.form_check = {'check_work_place': check_work_place, 'check_cronos': check_cronos,
                            'check_cross': check_cross, 'check_passport': check_passport,
                            'check_debt': check_debt, 'check_bankruptcy': check_bankruptcy,
                            'check_bki': check_bki, 'check_affiliation': check_affiliation,
                            'check_internet': check_internet, 'resume': resume,
                            'date_check': date_check, 'officer': officer
                            }
        return self.form_check

    # данные запроса на вставку данных кандидатов
    def cand_values_form(self):
        self.cand_values = tuple(map(str, [self.form_view['staff'], self.form_view['department'],
                                    self.form_view['full_name'], self.form_view['last_name'],
                                    self.form_view['birthday'], self.form_view['birth_place'],
                                    self.form_view['country'], self.form_view['series_passport'],
                                    self.form_view['number_passport'], self.form_view['date_given'],
                                    self.form_view['snils'], self.form_view['inn'],
                                    self.form_view['reg_address'], self.form_view['live_address'],
                                    self.form_view['phone'], self.form_view['email'],
                                    self.form_view['education'], self.form_check['check_work_place'],
                                    self.form_check['check_passport'], self.form_check['check_debt'],
                                    self.form_check['check_bankruptcy'], self.form_check['check_bki'],
                                    self.form_check['check_affiliation'], self.form_check['check_internet'],
                                    self.form_check['check_cronos'], self.form_check['check_cross'],
                                    self.form_check['resume'], self.form_check['date_check'],
                                    self.form_check['officer']]
                                     )
                                 )
        return self.cand_values

            
# запуск программы
if __name__ == "__main__":
    """Create backup"""
    # Создаем экземпляры класса BackUp и резервные копии MAIN_FILE, CONNECT
    main_file_copy, database_copy = BackUp(MAIN_FILE, DESTINATION), BackUp(CONNECT, DESTINATION)
    main_file_copy.backup()
    database_copy.backup()
    # открываем первый лист книги MAIN_FILE для чтения и записи данных
    wb = ExcelFile(MAIN_FILE)
    ws = wb.open_sheet(0)
    # Создаем список ячеек с датами согласования сегодня, ограничение 20 тыс. строк
    row_num = wb.range_row()
    # Создаем список ячеек с фамилией кандидата и убираем пробелы в начале и в конце
    fio = [ws['B' + str(i)].value.strip().lower() for i in row_num]
    # Перебираем каталоги в исходной папке
    search = DirSubdir(WORK_DIR)
    subdir = search.dir_range(fio)
    # Создаем список файлов Заключений, очищаем от пустых значений
    files = [search.file_range(f"{WORK_DIR[0:-1] + i}\\") for i in subdir]
    path_files = list(filter(None, files))
    form = Forms()  # создаем экземпляр класса для анкетных данных и заключений
    # открываем Заключения для чтения данных
    for path in path_files:
        wbz = ExcelFile(path)
        # проверяем количество листов в книге, если больше 1 и на 2-м листе есть данные
        if len(wbz.sheetnames) > 1 and str(wbz.worksheets[1]['K1'].value) == 'ФИО':
            wsz = wbz.open_sheet(1)
            form_view = form.view_form(wsz['C3'].value, wsz['D3'].value, wsz['K3'].value, wsz['S3'].value,
                                       wsz['L3'].value, wsz['M3'].value, wsz['T3'].value, wsz['P3'].value,
                                       wsz['Q3'].value, wsz['R3'].value, wsz['U3'].value, wsz['V3'].value,
                                       wsz['N3'].value, wsz['O3'].value, wsz['Y3'].value, wsz['Z3'].value,
                                       wsz['X3'].value)
        # если лист с анкетой отсутствует или на листе нет данных берем данные из заключения
        elif len(wbz.sheetnames) < 2 or (len(wbz.sheetnames) > 1 and
                                         str(wbz.worksheets[1]['K1'].value) != 'ФИО'):
            wsz = wbz.open_sheet(0)
            form_view = form.view_form(wsz['C4'].value, wsz['C5'].value, wsz['C6'].value, wsz['C7'].value,
                                       wsz['C8'].value, 'None', 'None', wsz['C9'].value, wsz['D9'].value,
                                       wsz['E9'].value, 'None', wsz['C10'].value, 'None', 'None', 'None', 
                                       'None', 'None')
        # Получаем данные из заключений и записываем значения
        wsz = wbz.open_sheet(0)
        form_check = form.check_form(f"{wsz['C11'].value} - {wsz['D11'].value}; {wsz['C12'].value} - "
                                    f"{wsz['D12'].value}; {wsz['C13'].value} - {wsz['D13'].value}",
                                    f"{wsz['B14'].value}: {wsz['C14'].value}; {wsz['B15'].value}: {wsz['C15'].value}",
                                    wsz['C16'].value, wsz['C17'].value, wsz['C18'].value, wsz['C19'].value,
                                     wsz['C20'].value, wsz['C21'].value, wsz['C22'].value, wsz['C23'].value,
                                     wsz['C24'].value, wsz['C25'].value)
        # Закрываем книгу Excel
        wbz.close_workbook()
        # создаем переменные с запросом и со значениями
        ins_cand = "INSERT INTO candidates ({}) VALUES ({}?)".format(SQL_CAND, '?, ' * (len(SQL_CAND.split()) - 1))
        val_cand = form.cand_values_form()
        # создаем экземпляр класса Database и передаем данные в БД
        candidate = Database(CONNECT, ins_cand, val_cand)
        candidate.insert_db()

    # Создание гиперссылок и перемещение папок
    links = HyperLink()
    hlink = links.create_link()
    # переносим папку из исходной в целевую папку
    for k, v in hlink.items():
        try:
            work_directory = BackUp(WORK_DIR + k, v)
            work_directory.remove()
        except shutil.Error:
            print('Ошибка при переносе данных')

    # Получаем данные из реестра кандидатов и формируем запрос в БД
    ins_reg = "INSERT INTO registry ({}) VALUES ({}?)".format(SQL_REG, '?, ' * (len(SQL_REG.split()) - 1))
    val_reg = wb.reg_range()
    # создаем экземпляр класса Database и передаем данные в БД
    reg = Database(CONNECT, ins_reg, val_reg)
    reg.insert_db()

    # Сохраняем книгу Excel
    wb.save_workbook(MAIN_FILE)
