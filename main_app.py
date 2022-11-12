import shutil
import sqlite3
from datetime import date
import os

import openpyxl

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
# CONNECT = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\candidates.db'
CONNECT = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\candidates.db'
# главный файл кандидатов
# MAIN_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'
MAIN_FILE = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'
# рабочая папка кандидатов
# WORK_DIR = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\\'
WORK_DIR = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\\'
# место хранения кандидатов
# DESTINATION = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Персонал\Персонал-2\\'
DESTINATION = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Персонал\Персонал-2\\'

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


# главный модуль программы
if __name__ == "__main__":

    """Создаем резервные копии"""
    # Создаем резервную копию книги Exel в DESTINATION
    shutil.copy(MAIN_FILE, DESTINATION)
    # Создаем резервную копию БД в DESTINATION
    shutil.copy(CONNECT, DESTINATION)

    """Работаем с файлом реестра кандидатов"""
    # открываем книгу MAIN_FILE для чтения и записи данных
    wb = openpyxl.load_workbook(MAIN_FILE, keep_vba=True)
    # Берем первый лист
    ws = wb.worksheets[0]
    # Идет поиск ячеек с датами согласования сегодня, ограничение 30 тыс. строк
    col_range = ws['K1':'K30000']
    for cell in col_range:
        for c in cell:
            # берем номер строки
            row_num = c.row
            # главное условие проверки - запись в ячейке равна сегодняшней дате
            if str(c.value) == DATE_TODAY:
                """ Получаем данные из заключений и формируем из них запрос в БД"""
                # берем значение с фамилией
                fio = ws['B' + str(row_num)].value
                # Перебираем каталоги в исходной папке
                for subdir in os.listdir(WORK_DIR):
                    # если имя папки такое же как и значение в ячейке фамилия
                    if subdir.lower().rstrip() == fio.lower().rstrip():
                        # ищем в папке файлы Заключение
                        for file in os.listdir(f"{WORK_DIR[0:-1] + subdir}\\"):
                            if file.startswith("Заключение"):
                                # открываем Заключение для чтения данных
                                wbz = openpyxl.load_workbook(os.path.join(WORK_DIR[0:-2], subdir, file), keep_vba=True)

                                """Получаем анкетные данные"""
                                form_dict = {}
                                # проверяем количество листов в книге, если больше 1 и на 2-м листе есть данные
                                if len(wbz.sheetnames) > 1 and str(wbz.worksheets[1]['K1'].value) == 'ФИО':
                                    # Берем второй лист
                                    ws1 = wbz.worksheets[1]
                                    # записываем значения
                                    form_dict = {'staff': ws1['C3'].value,
                                                 'department': ws1['D3'].value,
                                                 'full_name': ws1['K3'].value,
                                                 'last_name': ws1['S3'].value,
                                                 'birthday': ws1['L3'].value,
                                                 'birth_place': ws1['M3'].value,
                                                 'country': ws1['T3'].value,
                                                 'series_passport': ws1['P3'].value,
                                                 'number_passport': ws1['Q3'].value,
                                                 'date_given': ws1['R3'].value,
                                                 'snils': ws1['U3'].value,
                                                 'inn': ws1['V3'].value,
                                                 'reg_address': ws1['N3'].value,
                                                 'live_address': ws1['O3'].value,
                                                 'phone': ws1['Y3'].value,
                                                 'email': ws1['Z3'].value,
                                                 'education': ws1['X3'].value
                                                 }
                                # если лист с анкетой отсутствует или на листе нет данных берем данные из заключения
                                elif len(wbz.sheetnames) < 2 or (len(wbz.sheetnames) > 1 and
                                                                 str(wbz.worksheets[1]['K1'].value) != 'ФИО'):
                                    ws2 = wbz.worksheets[0]
                                    form_dict = {'staff': ws2['C4'].value,
                                                 'department': ws2['C5'].value,
                                                 'full_name': ws2['C6'].value,
                                                 'last_name': ws2['C7'].value,
                                                 'birthday': ws2['C8'].value,
                                                 'birth_place': 'None',
                                                 'country': 'None',
                                                 'series_passport': ws2['C9'].value,
                                                 'number_passport': ws2['D9'].value,
                                                 'date_given': ws2['E9'].value,
                                                 'snils': 'None',
                                                 'inn': ws2['C10'].value,
                                                 'reg_address': 'None',
                                                 'live_address': 'None',
                                                 'phone': 'None',
                                                 'email': 'None',
                                                 'education': 'None'
                                                 }

                                """Получаем данные проверки"""
                                # открываем первый лист с заключением
                                ws2 = wbz.worksheets[0]
                                # записываем значения
                                check_dict = {
                                    'check_work_place':
                                        str(ws2['C11'].value) + ' - ' +
                                        str(ws2['D11'].value) + '; ' +
                                        str(ws2['C12'].value) + ' - ' +
                                        str(ws2['D12'].value) + '; ' +
                                        str(ws2['C13'].value) + ' - ' +
                                        str(ws2['D13'].value),
                                    'check_cronos':
                                        str(ws2['B14'].value) + ': ' +
                                        str(ws2['C14'].value) + '; ' +
                                        str(ws2['B15'].value) + ': ' +
                                        str(ws2['C15'].value),
                                    'check_cross': ws2['C16'].value,
                                    'check_passport': ws2['C17'].value,
                                    'check_debt': ws2['C18'].value,
                                    'check_bankruptcy': ws2['C19'].value,
                                    'check_bki': ws2['C20'].value,
                                    'check_affiliation': ws2['C21'].value,
                                    'check_internet': ws2['C22'].value,
                                    'resume': ws2['C23'].value,
                                    'date_check': ws2['C24'].value,
                                    'officer': ws2['C25'].value,
                                    'check_rand_info': ws2['C29'].value
                                }
                                # Закрываем книгу Excel
                                wbz.close()

                                # создаем переменную с запросом
                                insert = f"INSERT INTO candidates ({SQL_CAND}) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " \
                                         f"?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                                # создаем переменную со значениями
                                values = tuple([str(i) for i in
                                                [form_dict['staff'], form_dict['department'], form_dict['full_name'],
                                                 form_dict['last_name'], form_dict['birthday'],
                                                 form_dict['birth_place'], form_dict['country'],
                                                 form_dict['series_passport'], form_dict['number_passport'],
                                                 form_dict['date_given'], form_dict['snils'], form_dict['inn'],
                                                 form_dict['reg_address'], form_dict['live_address'],
                                                 form_dict['phone'], form_dict['email'], form_dict['education'],
                                                 check_dict['check_work_place'], check_dict['check_passport'],
                                                 check_dict['check_debt'], check_dict['check_bankruptcy'],
                                                 check_dict['check_bki'], check_dict['check_affiliation'],
                                                 check_dict['check_internet'], check_dict['check_cronos'],
                                                 check_dict['check_cross'], check_dict['check_rand_info'],
                                                 check_dict['resume'], check_dict['date_check'],
                                                 check_dict['officer']
                                                 ]
                                                ]
                                               )
                                # создаем экземпляр класса Database
                                candidate = Database(CONNECT, insert, values)
                                # передаем данные в БД
                                candidate.insert_db()

                        """Переносим папки в архив"""
                        # разбираем посимвольно фамилию == имя папки
                        letter = [i for i in fio]
                        # создаем ссылку для DESTINATION с первым именем буквы по алфавиту и добавляем ID
                        hlink = DESTINATION + '\\' + letter[0] + '\\' + subdir + ' - ' + str(ws['A' +
                                                                                                str(row_num)].value)
                        # добавляем в файл книги гиперссылку, куда будет помещена папка
                        ws['L' + str(row_num)].hyperlink = hlink
                        # переносим папку из исходной в целевую папку
                        try:
                            shutil.move(WORK_DIR + subdir, hlink)
                        except shutil.Error:
                            continue
                """Получаем данные из реестра кандидатов и формируем запрос в БД"""
                # берем значения из ячеек, которые соответствуют номеру строки с сегодняшней датой
                reg_dict = {'fio': ws['B' + str(row_num)].value, 'birthday': ws['C' + str(row_num)].value,
                            'staff': ws['D' + str(row_num)].value, 'checks': ws['E' + str(row_num)].value,
                            'recruiter': ws['F' + str(row_num)].value, 'date_in': ws['G' + str(row_num)].value,
                            'officer': ws['H' + str(row_num)].value, 'date_out': ws['I' + str(row_num)].value,
                            'result': ws['J' + str(row_num)].value, 'fin_date': ws['K' + str(row_num)].value,
                            'url': ws['L' + str(row_num)].value
                            }
                # формируем запрос в таблицу реестр БД
                insert = f"INSERT INTO registry ({SQL_REG}) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                values = tuple([str(i) for i in
                                [reg_dict['fio'], reg_dict['birthday'], reg_dict['staff'], reg_dict['checks'],
                                 reg_dict['recruiter'], reg_dict['date_in'], reg_dict['officer'], reg_dict['date_out'],
                                 reg_dict['result'], reg_dict['fin_date'], reg_dict['url']
                                 ]
                                ]
                               )
                # создаем экземпляр класса Database
                reg = Database(CONNECT, insert, values)
                # передаем данные в БД
                reg.insert_db()

    # Сохраняем книгу Excel
    wb.save(MAIN_FILE)
