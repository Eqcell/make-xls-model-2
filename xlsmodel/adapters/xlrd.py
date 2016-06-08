import xlrd

from xlsmodel.adapters.base import ExcelReadObjectsFactory, Workbook, Sheet, Cell


class XlrdCell(Cell):

    def __init__(self, cell):
        self._cell = cell

    def get_value(self):
        return self._cell.value


class XlrdSheet(Sheet):

    def __init__(self, sheet):
        self._sheet = sheet

    def get_columns_number(self):
        return self._sheet.ncols

    def get_rows_number(self):
        return self._sheet.nrows

    def get_cell(self, row, column):
        return XlrdCell(self._sheet.cell(row, column))


class XlrdWorkbook(Workbook):

    def __init__(self, filename=None):
        self._workbook = None
        self.filename = None
        if filename:
            self.read_from_file(filename)

    def read_from_file(self, filename):
        with open(filename, 'rb') as f:
            self._workbook = xlrd.open_workbook(file_contents=f.read())
        self.filename = filename

    def get_sheet(self, index_or_name):
        """Return sheet by 1-based index or name"""
        try:
            index = int(index_or_name)
        except ValueError:
            name = index_or_name
            if name in self._workbook.sheet_names():
                sheet = self._workbook.sheet_by_name(name)
            else:
                raise ValueError('Failed to locate sheet: {}'.format(name))
        else:
            # TODO(dmu) LOW: What if index exceeds number of existing sheets?
            sheet = self._workbook.sheet_by_index(index - 1)  # because sheets are 0-based

        return XlrdSheet(sheet)


class ExcelXlrdReadObjectsFactory(ExcelReadObjectsFactory):

    def get_workbook(self, filename):
        return XlrdWorkbook(filename)
