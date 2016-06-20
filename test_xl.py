from basefunc import is_equal
from xl import ExcelSheet, get_xlrd_sheet
from test_model import DF, REF_DF, VAR_TO_ROWS, EQS
 
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

    assert SHEET_NAME == get_xlrd_sheet(PATH, 1).name
    assert SHEET_NAME == get_xlrd_sheet(PATH, sh_name).name
   
def test_xl_sheet_end_to_end():        
    
    ExcelSheet(PATH, 1, "A1").save(sheet=3)
    ExcelSheet(PATH, 2, "B3").save(sheet=4)

    df3 = ExcelSheet(PATH, sheet=3, anchor="A1").dataset
    df4 = ExcelSheet(PATH, sheet=4, anchor="B3").dataset  
    
    assert is_equal(df3, df4)   