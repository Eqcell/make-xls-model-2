import os
from xlmodel import ExcelSheet

ExcelSheet("test0.xls").save().echo()
ExcelSheet(os.path.join('examples', 'bdrn.xls'), "model", "c1").save().echo()