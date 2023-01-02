import openpyxl
import re


class Forms:
    """Объявляем класс Forms для работы с данными"""

    resumes: dict
    checks: dict

    def __init__(self, path):
        self.book = openpyxl.load_workbook(path, keep_vba=True)
        self.sheet = None
        self.check_excel()

    def check_excel(self):  # parse Excel files with conclusions and resume dates
        self.sheet = self.book.worksheets[0]
        self.get_check()
        self.get_conclusion_resume()
        if len(self.book.sheetnames) > 1:
            sheet = self.book.worksheets[1]
            if str(sheet['K1'].value) == 'ФИО':
                self.sheet = sheet
                self.get_full_resume()
        self.book.close()

    @staticmethod
    def data_check(val):    # check string format date
        str_date = None
        if isinstance(val, str):
            data = re.findall(r'\d{2}.\d{2}.\d{4}', val)
            if len(data):
                view = ' '.join(data).strip()
                str_date = f'{view[-4:]}-{view[-7:-5]}-{view[:3]}'
            else:
                str_date = val.strip()
        return str_date

    def get_full_resume(self):
        Forms.resumes = {'staff': self.sheet['C3'].value,
                         'department': self.sheet['D3'].value,
                         'full_name': self.sheet['K3'].value,
                         'last_name': self.sheet['S3'].value,
                         'birthday': Forms.data_check(self.sheet['L3'].value),
                         'birth_place': self.sheet['M3'].value,
                         'country': self.sheet['T3'].value,
                         'series_passport': self.sheet['P3'].value,
                         'number_passport': self.sheet['Q3'].value,
                         'date_given': Forms.data_check(self.sheet['R3'].value),
                         'snils': self.sheet['U3'].value,
                         'inn': self.sheet['V3'].value,
                         'reg_address': self.sheet['N3'].value,
                         'live_address': self.sheet['O3'].value,
                         'phone': self.sheet['Y3'].value,
                         'email': self.sheet['Z3'].value,
                         'education': self.sheet['X3'].value
                         }
        return Forms.resumes

    def get_conclusion_resume(self):
        Forms.resumes = {'staff': self.sheet['C4'].value,
                         'department': self.sheet['C5'].value,
                         'full_name': self.sheet['C6'].value,
                         'last_name': self.sheet['C7'].value,
                         'birthday': Forms.data_check(self.sheet['C8'].value),
                         'birth_place': '',
                         'country': '',
                         'series_passport': self.sheet['C9'].value,
                         'number_passport': self.sheet['D9'].value,
                         'date_given': Forms.data_check(self.sheet['E9'].value),
                         'snils': '',
                         'inn': self.sheet['C10'].value,
                         'reg_address': '',
                         'live_address': '',
                         'phone': '',
                         'email': '',
                         'education': ''
                         }
        return Forms.resumes

    def get_check(self):
        Forms.checks = {'check_work_place':
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
                        'date_check': Forms.data_check(self.sheet['C24'].value),
                        'officer': self.sheet['C25'].value,
                        }
        return Forms.checks


class Registries:
    """Class for registry and inquiry dates"""

    def __init__(self, sheet, num):
        self.sheet = sheet
        self.num = num
        self.registry = None
        self.inquiry = None

    def get_registry(self):
        self.registry = {'checks': self.sheet[f'E{self.num}'].value,
                         'recruiter': self.sheet[f'F{self.num}'].value,
                         'final_date': self.sheet[f'K{self.num}'].value,
                         'url': Forms.data_check(self.sheet[f'L{self.num}'].value)}
        return self.registry

    def get_inquiry(self):
        self.inquiry = {'staff': self.sheet[f'C{self.num}'].value,
                        'period': self.sheet[f'D{self.num}'].value,
                        'info': self.sheet[f'E{self.num}'].value,
                        'firm': self.sheet[f'F{self.num}'].value,
                        'date_inq': Forms.data_check(self.sheet[f'G{self.num}'].value)}
        return self.inquiry
