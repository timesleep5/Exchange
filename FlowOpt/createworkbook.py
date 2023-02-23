from typing import List

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


class CreateWorkbook:
    def __init__(self, wbname):
        self.wb = self._get_wb()
        self.week = ['Monday', 'Tuesday', 'Wednesday', 'Tuesday', 'Friday', 'Saturday', 'Sunday']
        self.categories = ['Day', 'Date', 'Planned', 'Done', 'Resumé']
        self.wbname = wbname
        self.create()

    @staticmethod
    def _get_wb():
        return Workbook()

    def get_workheets(self):
        self.wb.create_sheet()
        self.wb.create_sheet()

    def style_sheets(self):
        self.style_backlog()
        self.style_weektable()
        self.style_history()

    def style_backlog(self):
        ws = self.wb['Sheet']
        ws.title = 'Backlog'
        ws.append(self.week)
        self._make_row_bold(ws, 1, len(self.week), 1)

    def style_weektable(self):
        ws = self.wb['Sheet1']
        ws.title = 'WeekTable'
        ws.append(self.week)
        self._make_row_bold(ws, 1, len(self.week), 1)

    def style_history(self):
        ws = self.wb['Sheet2']
        ws.title = 'History'
        for i in range(1, len(self.categories) + 1):
            ws[f'A{i}'] = self.categories[i - 1]
        self._fill_row(ws, 2, len(self.week)+1, 1, self.week)
        self._fill_col(ws, 1, len(self.categories), 1, self.categories)
        self._make_col_bold(ws, 1, len(self.categories), 'A')

    def _fill_row(self, ws, startcol: int, endcol: int, row: int, content: List[str]):
        for i in range(startcol, endcol + 1):
            char = get_column_letter(i)
            ws[f'{char}{row}'] = content[i - startcol] # to be at Index 0 in the first iteration

    def _fill_col(self, ws, startrow:int, endrow: int, col: int, content: List[str]):
        char = get_column_letter(col)
        for i in range(startrow, endrow+1):
            ws[f'{char}{i}'] = content[i-startrow]

    def _make_row_bold(self, ws, startcol, endcol, row):
        for i in range(startcol, endcol+1):
            char = get_column_letter(i)
            ws[f'{char}{row}'].font = Font(bold=True)

    def _make_col_bold(self, ws, startrow, endrow, col):
        for i in range(startrow, endrow+1):
            ws[f'{col}{i}'].font = Font(bold=True)

    def create(self):
        self.get_workheets()
        self.style_sheets()
        self.savewb()

    def savewb(self):
        self.wb.save(filename=self.wbname)