import openpyxl
from datetime import datetime


class Forms:
    """Объявляем класс Forms для работы с данными"""

    def __init__(self, path):
        self.book = openpyxl.load_workbook(path, keep_vba=True)
        self.sheet = None
        self.resume = None
        self.check = None
        self.concluded = None
        self.check_excel()

    def check_excel(self): # parse Excel files with conclusions and resume dates
        self.sheet = self.book.worksheets[0]
        self.get_check()
        self.get_conclusion_resume()
        if len(self.book.sheetnames) > 1:
            sheet = self.book.worksheets[1]
            if str(sheet['K1'].value) == 'ФИО':
                self.sheet = sheet
                self.get_full_resume()
        self.book.close()


    def get_full_resume(self):
        self.resume = {'staff': self.sheet['C3'].value,
                     'department': self.sheet['D3'].value,
                     'full_name': self.sheet['K3'].value,
                     'last_name': self.sheet['S3'].value,
                     'birthday':  datetime.strptime(self.sheet['L3'].value, '%d.%M.%Y'),
                     'birth_place': self.sheet['M3'].value,
                     'country': self.sheet['T3'].value,
                     'series_passport': self.sheet['P3'].value,
                     'number_passport': self.sheet['Q3'].value,
                     'date_given': datetime.strptime(self.sheet['R3'].value, '%d.%M.%Y'),
                     'snils': self.sheet['U3'].value,
                     'inn': self.sheet['V3'].value,
                     'reg_address': self.sheet['N3'].value,
                     'live_address': self.sheet['O3'].value,
                     'phone': self.sheet['Y3'].value,
                     'email': self.sheet['Z3'].value,
                     'education': self.sheet['X3'].value
                     }
        return self.resume

    def get_conclusion_resume(self):
        self.resume = {'staff': self.sheet['C4'].value,
                     'department': self.sheet['C5'].value,
                     'full_name': self.sheet['C6'].value,
                     'last_name': self.sheet['C7'].value,
                     'birthday': self.sheet['C8'].value,
                     'birth_place': 'None',
                     'country': 'None',
                     'series_passport': self.sheet['C9'].value,
                     'number_passport': self.sheet['D9'].value,
                     'date_given': self.sheet['E9'].value,
                     'snils': 'None',
                     'inn': self.sheet['C10'].value,
                     'reg_address': 'None',
                     'live_address': 'None',
                     'phone': 'None',
                     'email': 'None',
                     'education': 'None'
                     }
        return self.resume

    def get_check(self):
        self.check = {'check_work_place':
                        f"{self.sheet['C11'].value} - {self.sheet['D11'].value}; {self.sheet['C12'].value} - "
                        f"{self.sheet['D12'].value}; {self.sheet['C13'].value} - {self.sheet['D13'].value}",
                    'check_cronos':
                        f"{self.sheet['B14'].value}: {self.sheet['C14'].value}; {self.sheet['B15'].value}: "
                        f"{self.sheet['C15'].value}",
                    'check_cross': self.sheet['C16'].value,
                    'check_passport': self.sheet['C17'].value,
                    'check_debt': self.sheet['C18'].value,
                    'check_bankruptcy': self.sheet['C19'].value,
                    'check_bki': self.sheet['C20'].value,
                    'check_affiliation': self.sheet['C21'].value,
                    'check_internet': self.sheet['C22'].value,
                    'resume': self.sheet['C23'].value,
                    'date_check': self.sheet['C24'].value,
                    'officer': self.sheet['C25'].value,
                    }
        return self.check

    def get_conclusion(self):
        self.concluded = [self.resume['staff'],
                          self.resume['department'],
                          self.resume['full_name'],
                          self.resume['last_name'],
                          self.resume['birthday'],
                          self.resume['birth_place'],
                          self.resume['country'],
                          self.resume['series_passport'],
                          self.resume['number_passport'],
                          self.resume['date_given'],
                          self.resume['snils'],
                          self.resume['inn'],
                          self.resume['reg_address'],
                          self.resume['live_address'],
                          self.resume['phone'],
                          self.resume['email'],
                          self.resume['education'],
                          self.check['check_work_place'],
                          self.check['check_passport'],
                          self.check['check_debt'],
                          self.check['check_bankruptcy'],
                          self.check['check_bki'],
                          self.check['check_affiliation'],
                          self.check['check_internet'],
                          self.check['check_cronos'],
                          self.check['check_cross'],
                          self.check['resume'],
                          self.check['date_check'],
                          self.check['officer']
                        ]
        return self.concluded