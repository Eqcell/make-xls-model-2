import os
import pandas as pd
import numpy as np

from xlmodel import col_to_num, to_xl_ref, to_rowcol
from xlmodel import FormulaSegment, Formula, MathModel  
from xlmodel import ExcelSheet, _get_xlrd_sheet

from xlmodel import is_equal


def test_is_equal():
    # not tested
    pass
 
def test_basefunc():
    # Excel references
    assert col_to_num("A") == 1
    assert col_to_num("B") == 2
    assert to_xl_ref(1, 1) == "A1"
    assert to_xl_ref(1, 1, base = 1) == "A1"
    assert to_xl_ref(0, 0, base = 0) == "A1"    
    assert to_rowcol("A1") == (1, 1)
    assert to_rowcol("A1", base = 0) == (0, 0)
    assert to_rowcol("AA1") == (1, 27)
    
   
# tests model.py and some of xl.py

# test data     
COLUMNS = ['is_forecast', 'y', 'rog']   
VAR_TO_ROWS = {'is_forecast': 2, 'y' : 3, 'rog' : 4}
DF =  pd.DataFrame({  'y' : [    85,    100, np.nan],
                    'rog' : [np.nan, np.nan,   1.05],
            'is_forecast' : [     0,      0,      1]},
                    index = [  2014,   2015,   2016])[COLUMNS]   
assert is_equal(DF, pd.read_excel('test1.xls').transpose()[COLUMNS])
EQS = ['y = y[t-1] * rog'] 
REF_DF = DF.copy()
REF_DF.loc[2016,'y'] = '=C3*D4'

def test_segment():
    # (varname + time period) segment 
    # test segment "GDP[1]" conversion to 'B5' 
    fs = FormulaSegment("GDP[1]", {'GDP':5}, anchor = "A1")    
    assert fs.col == 1
    assert fs.row == 5
    assert fs.column_offset == 1
    assert fs.xl_ref() == 'B5'

def test_formula(): 
    # formula strings converted to xl references
    # common arguments    
    pos = VAR_TO_ROWS, "A1"    
    # time period 1 + column offset 1 = B     
    assert "=B2" == Formula('is_forecast[t]', *pos).get_xl_formula(time_period=1)
    # time period 3 + column offset 1 = D
    assert '=C3*D4' == Formula('y[t-1] * rog', *pos).get_xl_formula(time_period=3) 
    # whitespace stripped
    assert "GDP[t]" == Formula("  GDP[t]  ", *pos).__repr__()
    # same start of variable name
    assert 'FondOT[t]+FondOther[t]' == Formula.expand_shorthand("FondOT+FondOther", {"FondOT":0,"FondOther":1})
    
def test_math_model():
    # model with no Excel, local variables only
    m = MathModel(equations = EQS, dataset = DF)
    m.set_xl_positioning(var_to_rows = VAR_TO_ROWS)
    assert is_equal(m.get_xl_dataset(), REF_DF)
    
PATH = "test1.xls"
SHEET_NAME = 'input_sheet_v1'
   
def test_xl_sheet_reading():    
    
    df1 = ExcelSheet(PATH, sheet=1, anchor="A1").dataset
    df2 = ExcelSheet(PATH, sheet=2, anchor="B3").dataset    
    assert is_equal(df2, df1)

def test_model_on_sheet():
    sh = ExcelSheet(PATH)
    assert is_equal(sh.dataset, DF)
    assert sh.var_to_rows == VAR_TO_ROWS
    assert sh.equations == EQS
    assert sh.model.equations['y'] == 'y[t-1] * rog'
    assert is_equal(REF_DF, sh.model.get_xl_dataset())    
    
def read_by_name_and_int():

    assert SHEET_NAME == _get_xlrd_sheet(PATH, 1).name
    assert SHEET_NAME == _get_xlrd_sheet(PATH, sh_name).name
   
def test_xl_sheet_end_to_end():        
    
    ExcelSheet(PATH, 1, "A1").save(sheet=3)
    ExcelSheet(PATH, 2, "B3").save(sheet=4)

    df3 = ExcelSheet(PATH, sheet=3, anchor="A1").dataset
    df4 = ExcelSheet(PATH, sheet=4, anchor="B3").dataset  
    
    assert is_equal(df3, df4)

def run_example(filename, sheet=1, anchor="c1"):
    ExcelSheet(os.path.join('examples', filename), sheet, anchor).save()#.echo()
    
def test_examples_folder():

    ExcelSheet("test0.xls").save()
    ExcelSheet("test1.xls", 1, "A1").save(sheet=3)
    ExcelSheet("test1.xls", 2, "B3").save(sheet=4)

    run_example('bdrn.xls')
    run_example('ref_file.xls')
    run_example('spec.xls')
    run_example('spec2.xls')
    run_example('bank.xls')
    run_example('bank_sector.xls')