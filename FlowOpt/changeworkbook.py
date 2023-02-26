from datetime import date

from openpyxl import load_workbook

from workbookutils import Workbookutils as w_utils

from timeutils import Timeutils as t_utils

class ChangeWorkbook:
    def __init__(self, wbname):
        self.wb = self.get_workbook(wbname)

    @staticmethod
    def get_workbook(wbname):
        return load_workbook(wbname)

    def update_schedule_dates(self):
        ws = self.wb['History']
        dates = t_utils.get_week_dates(date.today(), 1, 7)
        w_utils.fill_row(ws, 2, len(dates)+1, 2, dates) #todo

    def add_to_backlog(self):
        pass

    def remove_from_backlog(self):
        pass

    def move_from_backlog_to_schedule(self):
        pass

    def move_form_schedule_to_backlog(self):
        pass

