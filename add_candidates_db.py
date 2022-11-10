import os

from class_constant import Database, ExelFile, MAIN_FILE, CONNECT, DATE_TODAY, WORK_DIR, SQL_CAND


def candidates_add():
    # Создаем экземпляр класса ExcelFile
    # Открываем книгу по адресу MAIN_FILE для чтения и записи данных
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
                        # ищем в папке файлы Заключение
                        for file in os.listdir(f"{WORK_DIR[0:-1] + subdir}\\"):
                            if file.startswith("Заключение"):
                                # Создаем экземпляр класса ExcelFile
                                # Открываем книгу заключение для чтения и записи данных
                                wbz = ExelFile(os.path.join(WORK_DIR[0:-2], subdir, file))
                                # Берем второй лист
                                try:
                                    ws1 = wbz.open_file(1)
                                    # проверяем имеются ли данные на листе с анкетой
                                    if str(ws1['K1'].value) == 'ФИО':
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
                                    # если анкета отсутствует, берем анкетные данные с листа заключение
                                    else:
                                        ws2 = wbz.open_file(0)
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
                                                     'reg_address': 'NULL',
                                                     'live_address': 'NULL',
                                                     'phone': 'NULL',
                                                     'email': 'NULL',
                                                     'education': 'NULL'
                                                     }
                                # если отсутствует лист с анкетой то выполняем предыдущий блок
                                except IndexError:
                                    ws2 = wbz.open_file(0)
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
                                                 'reg_address': 'NULL',
                                                 'live_address': 'NULL',
                                                 'phone': 'NULL',
                                                 'email': 'NULL',
                                                 'education': 'NULL'
                                                 }

                                # открываем первый лист с заключением
                                ws2 = wbz.open_file(0)
                                # записываем значения
                                check_dict = {'check_work_place': str(ws2['C11'].value) +
                                                                  ' - ' +
                                                                  str(ws2['D11'].value) +
                                                                  '; ' +
                                                                  str(ws2['C12'].value) +
                                                                  ' - ' +
                                                                  str(ws2['D12'].value) +
                                                                  '; ' +
                                                                  str(ws2['C13'].value) +
                                                                  ' - ' +
                                                                  str(ws2['D13'].value),
                                              'check_cronos': str(ws2['B14'].value) +
                                                              ': ' +
                                                              str(ws2['C14'].value) +
                                                              '; ' +
                                                              str(ws2['B15'].value) +
                                                              ': ' +
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
                                # создаем переменную с запросом
                                insert = f"INSERT INTO candidates ({SQL_CAND}) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " \
                                         f"?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                                # создаем переменную с запросом
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
                                                 check_dict['officer']]])
                                # создаем экземпляр класса Database
                                candidate = Database(CONNECT, insert, values)
                                # передаем данные в БД
                                candidate.insert_db()

    # Сохраняем книгу Excel
    wb.save(MAIN_FILE)
