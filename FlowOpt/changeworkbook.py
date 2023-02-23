from openpyxl import Workbook

class ChangeWorkbook:
    def __init__(self, wb):
        self.wb = wb


    def test(self):
        ws = self.wb.active
        print(self.ws['A1'])