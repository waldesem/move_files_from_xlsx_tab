import shutil
import os
from datetime import date, datetime

from conclude import *
from database import *

DATE = date.today()  # сегодняшняя дата (объект)
# CONNECT = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\candidates.db' #     файл базы данных
# WORK_DIR = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\\'    # рабочая папка кандидатов
# DESTINATION = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Персонал\Персонал-2\\'   # архивная папка
# INFO_FILE = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\Запросы по работникам.xlsx'  # запросы
# MAIN_FILE = r'\\cronosx1\New folder\УВБ\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'  # файл кандидатов
# тестовые константы:
CONNECT = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\personal.db'  # файл базы данных
WORK_DIR = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\\'  # рабочая папка
DESTINATION = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Персонал\Персонал-2\\'  # архивная папка
MAIN_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Кандидаты.xlsm'  # файл кандидатов
INFO_FILE = r'C:\Users\ubuntu\Documents\Отдел корпоративной защиты\Кандидаты\Запросы по работникам.xlsx'  # запросы


def range_row(sheet):  # get list of row nums that correspond with today date
    row_num = []
    for cell in sheet:
        for c in cell:
            if isinstance(c.value, datetime):  # check format of date
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
            excel_to_db(path_files)  # parse files and send info to database
        create_link(sheet, subdirectory, num_row)  # create url and move folders to archive


def dir_range(sheet, row_num):  # получаем список папок в рабочей директории
    fio = [sheet['B' + str(i)].value.strip().lower() for i in row_num]
    subdir = [sub for sub in os.listdir(WORK_DIR) if sub.lower().strip() in fio]
    return subdir


def file_range(subdir):  # получаем список путей к файлам Заключений
    name_path = list(filter(None, []))
    for f in subdir:
        subdir_path = os.path.join(WORK_DIR[0:-1], f)
        for file in os.listdir(subdir_path):
            if file.startswith("Заключение") and (file.endswith("xlsm") or file.endswith("xlsx")):
                name_path.append(os.path.join(WORK_DIR, subdir_path, file))
    return name_path


def check_values(value):  # convert data types to string
    view = value.strftime('%Y-%m-%d') if isinstance(value, datetime) else str(value).strip()
    return view


def create_link(sheet, subdir, row_num):  # создаем гиперссылки и перемещаем папки
    for n in row_num:
        for sub in subdir:
            if str(sheet['B' + str(n)].value.strip().lower()) == sub.strip().lower():
                sbd = sheet['B' + str(n)].value.strip()
                lnk = os.path.join(DESTINATION, sbd[0][0], f"{sbd} - {sheet['A' + str(n)].value}")
                sheet['L' + str(n)].hyperlink = str(lnk)  # записывает в книгу
                # shutil.move(os.path.join(WORK_DIR, sbd), lnk)


def excel_to_db(path_files):  # take path's to conclusions
    for path in path_files:
        if len(path_files):
            form = Forms(path)  # take conclusions data
            fio, birthday = form.resumes['full_name'], check_values(form.resumes['birthday'])
            resume = {key: check_values(value) for key, value in form.resumes.items()}
            check = {key: check_values(value) for key, value in form.checks.items()}
            with Session(ENGINE) as sess:  # get personal dates
                result = sess.query(Candidate).filter_by(full_name=fio, birthday=birthday).first()
            if result is None:  # if no same data in db - add personal date and checks result
                value = Candidate(**resume)
                with Session(ENGINE) as sess:
                    sess.add(value)
                    sess.flush()
                    check['check_id'] = value.id
                    value = Check(**check)
                    sess.add(value)
                    sess.commit()
            else:   # if same data in db - update personal date and add checks result
                with Session(ENGINE) as sess:
                    search = sess.query(Candidate).get(result.id)
                    for k, v in resume.items():
                        setattr(search, k, v)
                    check['check_id'] = result.id
                    value = Check(**check)
                    sess.add(value)
                    sess.commit()


def chart_check(sheet, num_row, chart_id, chart):  # get data from registry an inquiry
    for num in num_row:
        reg = Registries(sheet, num)    
        if chart_id == 'registry_id': # get date from registry
            dictionary = {key: check_values(value) for key, value in reg.get_registry().items()}
            fio, birthday = sheet['B' + str(num)].value, check_values(sheet['C' + str(num)].value)
        else:   # get date from inquiry
            dictionary = {key: check_values(value) for key, value in reg.get_inquiry().items()}
            fio, birthday = sheet['A' + str(num)].value, check_values(sheet['B' + str(num)].value)
        with Session(ENGINE) as sess:   # check current values
            result = sess.query(Candidate).filter_by(full_name=fio, birthday=birthday).first()
        if result is None:  # if no same data, add fio and birth to Candidate table and others to Registr and Inquiry
            value = Candidate(**{'full_name': fio, 'birthday': birthday})
            with Session(ENGINE) as sess:
                sess.add(value)
                sess.flush()
                dictionary[chart_id] = value.id
                value = chart(**dictionary)
                sess.add(value)
                sess.commit()
        else:  # if same data id db is True, add value to Registr or Inquiry table
            dictionary[chart_id] = result.id
            value = chart(**dictionary)
            with Session(ENGINE) as sess:
                sess.add(value)
                sess.commit()
