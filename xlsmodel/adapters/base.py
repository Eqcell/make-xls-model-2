class ExcelObjectsFactory(object):
    def get_workbook(self, filename):
        raise NotImplementedError()


class ExcelReadObjectsFactory(ExcelObjectsFactory):
    pass


class ExcelWriteObjectsFactory(ExcelObjectsFactory):
    pass


class ExcelFactoriesFactory(object):

    def get_excel_read_objects_factory(self):
        from xlsmodel.adapters.xlrd import ExcelXlrdReadObjectsFactory
        return ExcelXlrdReadObjectsFactory()


class Cell(object):

    def get_value(self):
        raise NotImplementedError()


class Sheet(object):

    def get_columns_number(self):
        raise NotImplementedError()

    def get_rows_number(self):
        raise NotImplementedError()

    def get_cell(self, row, column):
        raise NotImplementedError()


class Workbook(object):

    def read_from_file(self, filename):
        raise NotImplementedError()

    def get_sheet(self, index_or_name):
        raise NotImplementedError()


excel_factories_factory = ExcelFactoriesFactory()
