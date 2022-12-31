from actions import *


class Main:
    """Initialize class for checking date of file changes"""

    def __init__(self):
        self.main_file_date = date.fromtimestamp(os.path.getmtime(MAIN_FILE))
        self.info_file_date = date.fromtimestamp(os.path.getmtime(INFO_FILE))
        self.check_date()

    def check_date(self):   # check file changing date
        if self.main_file_date == DATE or self.info_file_date == DATE:
            for file in (CONNECT, MAIN_FILE,  INFO_FILE):
                shutil.copy(file, DESTINATION)
        if self.main_file_date == DATE:
            check_main = Parse(MAIN_FILE, 'K5000', 'K15000')
            check_main.cand_check()
        if self.info_file_date == DATE:
            check_inq = Parse(INFO_FILE, 'G1', 'G2000')
            check_inq.inquiry_check()


class Parse:
    """Initialize class for parsing directories and files"""

    def __init__(self, file, start, end): # check today date in rows, create list with row numbers
        self.file = file
        self.wb = openpyxl.load_workbook(self.file, keep_vba=True, read_only=False)
        self.ws = self.wb.worksheets[0]
        self.num_row = range_row(self.ws[start:end])

    def cand_check(self):  # parse registry and conclusions
        if len(self.num_row):
            parse_conclusions(self.ws, self.num_row)
            registry_check(self.num_row, self.ws, 'B', 'L', 'registry', SQL_REG)
            self.wb.save(self.file)
        else:
            self.wb.close()
    def inquiry_check(self):  # take info from iquery registry and send it to database
        if len(self.num_row):
            registry_check(self.num_row, self.ws, 'A', 'G', 'inquiry', SQL_INQ)
        self.wb.close()


if __name__ == "__main__":
    Main()
