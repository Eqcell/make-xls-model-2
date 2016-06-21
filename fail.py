import os
from xlmodel import ExcelSheet
     
def example_full_path(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'examples', filename)
     
path = example_full_path('bdrn.xls')
ExcelSheet(path, "model", "c1").save().echo()
