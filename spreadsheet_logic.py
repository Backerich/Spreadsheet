import openpyxl


class Workbook(object):
    def get_workbook(self, workbook):
        return openpyxl.load_workbook(workbook)


class Sheets(object):
    def list_of_sheets(self, wb):
        return wb.get_sheet_names()

    def get_sheet(self, wb, sheet):
        # wenn Enter dannn:
        # return wb.get_active_sheet()
        return wb.get_sheet_by_name(sheet)