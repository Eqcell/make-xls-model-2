import os
from xlmodel import ExcelSheet

def run_example(filename, sheet=1, anchor="c1"):
    # as in ExcelSheet("test0.xls").save()
    ExcelSheet(os.path.join('examples', filename), sheet, anchor).save().echo()
    
run_example('bdrn.xls')
run_example('ref_file.xls')
run_example('spec.xls')
run_example('spec2.xls')
run_example('bank.xls')
run_example('bank_sector.xls')


