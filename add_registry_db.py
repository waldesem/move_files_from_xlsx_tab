from class_constant import Database, ExelFile, MAIN_FILE, CONNECT, DATE_TODAY, SQL_REG


def add_to_registry():
    # Создаем экземпляр класса ExcelFile и открываем книгу MAIN_FILE для чтения и записи данных
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
                values = str(reg_dict['fio']), str(reg_dict['birthday']), str(reg_dict['staff']), \
                         str(reg_dict['checks']), str(reg_dict['recruiter']), str(reg_dict['date_in']), \
                         str(reg_dict['officer']), str(reg_dict['date_out']), str(reg_dict['result']), \
                         str(reg_dict['fin_date']), str(reg_dict['url'])

                # создаем экземпляр класса Database
                reg = Database(CONNECT, insert, values)
                # передаем данные в БД
                reg.insert_db()

    # Сохранить книгу Excel
    wb.save(MAIN_FILE)
