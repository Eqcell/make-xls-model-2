# tests both xl.py and model.py

import pandas as pd
import numpy as np

from basefunc import is_equal
from model import FormulaSegment, Formula, MathModel  
from xl import ModelOnSheet

# test data     
COLUMNS = ['is_forecast', 'y', 'rog']   
VAR_TO_ROWS = {'is_forecast': 2, 'y' : 3, 'rog' : 4}
DF =  pd.DataFrame({  'y' : [    85,    100, np.nan],
                    'rog' : [np.nan, np.nan,   1.05],
            'is_forecast' : [     0,      0,      1]},
                    index = [  2014,   2015,   2016])[COLUMNS]   
assert is_equal(DF, pd.read_excel('xl.xls').transpose()[COLUMNS])
EQS = ['y = y[t-1] * rog'] 
REF_DF = DF.copy()
REF_DF.loc[2016,'y'] = '=C3*D4'

def test_segment():
    # segment 
    fs = FormulaSegment("GDP[1]", {'GDP':5}, anchor = "A1")    
    assert fs.col == 1
    assert fs.row == 5
    assert fs.offset == 1
    assert fs.xl_ref() == 'B5'

def test_formula(): 
    # formula strings converted to xl      
    pos = VAR_TO_ROWS, "A1"    
    # 1 + 1 = B     
    assert "=B2" == Formula('is_forecast[t]', *pos).get_xl_formula(1)
    # 3 + 1 = D
    assert '=C3*D4' == Formula('y[t-1] * rog', *pos).get_xl_formula(3) 
    # testing whitespace stripped
    assert "GDP[t]" == Formula("  GDP[t]  ", *pos).__repr__()
    
def test_math_model():
    # model with no Excel, local variables only
    m = MathModel(equations = EQS, dataset = DF)
    m.set_xl_positioning(var_to_rows = VAR_TO_ROWS)
    assert is_equal(m.get_xl_dataset(), REF_DF)

def test_to_matrix():
    # standalone method
    assert ModelOnSheet.to_matrix(DF) == [['', 'is_forecast', 'y', 'rog'], ['2014', '0', '85', ''], ['2015', '0', '100', ''], ['2016', '1', '', '1.05']]
   
def test_mos():
    mos = ModelOnSheet('xl.xls')
    assert is_equal(mos.dataset, DF)
    assert mos.var_to_rows == VAR_TO_ROWS
    assert mos.equations == EQS
    assert mos.model.equations['y'] == 'y[t-1] * rog'
    assert is_equal(REF_DF, mos.model.get_xl_dataset()) 