import shutil
import os

from class_constant import ExelFile, MAIN_FILE, DATE_TODAY, WORK_DIR, DESTINATION


def move_folder():
    # Создаем экземпляр класса ExcelFile и открываем книгу MAIN_FILE для чтения
    wb = ExelFile(MAIN_FILE)
    # Берем первый лист
    ws = wb.open_file(0)
    # Идет поиск ячеек с датами согласования сегодня, ограничение 30 тыс. строк
    col_range = ws['K1':'K30000']
    for cell in col_range:
        for c in cell:
            # берем номер строки
            row_num = c.row
            # если запись в ячейке равна сегодняшней дате
            if str(c.value) == DATE_TODAY:
                # берем значение с фамилией
                fio = ws['B' + str(row_num)].value
                # Перебираем каталоги в исходной папке
                for subdir in os.listdir(WORK_DIR):
                    # если имя папки такое же как и значение в ячейке фамилия
                    if subdir.lower().rstrip() == fio.lower().rstrip():
                        # разбираем посимвольно имя папки
                        letter = [i for i in fio]
                        # создаем ссылку для целевой директории с первым именем буквы по алфавиту
                        # добавляем к имени уникальный ID
                        hlink = DESTINATION + '\\' + letter[0] + '\\' + subdir + ' - ' + str(ws['A' +
                                                                                                str(row_num)].value)
                        # добавляем в файл книги гиперссылку, куда будет помещена папка
                        ws['L' + str(row_num)].hyperlink = hlink
                        # переносим папку из исходной в целевую папку
                        try:
                            shutil.move(WORK_DIR + subdir, hlink)
                        except shutil.Error:
                            continue
    # Сохраняем книгу Excel
    wb.save(MAIN_FILE)
