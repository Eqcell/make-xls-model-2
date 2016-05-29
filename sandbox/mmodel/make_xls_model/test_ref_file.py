import numpy as np 

import os 

from make_xl_model import get_resulting_workbook_array_for_make
from make_xl_model import get_array_and_support_variables, fill_array_with_excel_formulas_based_on_is_forecast

abs_filepath = os.path.abspath(os.path.join("examples","ref_file.xls"))
pivot_col = 2
sheet = 'model'

# expected array strings
make_string = """[['' '2010' '2011' '2012' '2013' '2014' '2015' '2016' '2017' '2018']
 ['x' 100.0 101.0 102.5 100.8 95.5 102.5 105.0 '=I4' '=J4']
 ['y' 3360.0 2700.0 500.0 1200.0 4800.0 5280.0 6336.0 '=H3*I5' '=I3*J5']
 ['x_fut' nan nan nan nan nan nan nan 102.5 105.0]
 ['y_rog' nan nan nan nan nan nan nan 1.1 1.2]
 ['is_forecast' 0.0 0.0 0.0 0.0 0.0 0.0 0.0 1.0 1.0]]"""

update_string = """[['' '' '' 2010.0 2011.0 2012.0 2013.0 2014.0 2015.0 2016.0 2017.0 2018.0]
 ['1. ИСХОДНЫЕ ДАННЫЕ И ПРОГНОЗ' '' '' '' '' '' '' '' '' '' '' '']
 ['' '' 'x' 100.0 101.0 102.5 100.8 95.5 102.5 105.0 '=K6' '=L6']
 ['' '' 'y' 3360.0 2700.0 500.0 1200.0 4800.0 5280.0 6336.0 '=J4*K7'
  '=K4*L7']
 ['2. УПРАВЛЯЮЩИЕ ПАРАМЕТРЫ' '' '' '' '' '' '' '' '' '' '' '']
 ['' 'ФУТ' 'x_fut' '' '' '' '' '' '' '' 102.5 105.0]
 ['' 'РОГ' 'y_rog' '' '' '' '' '' '' '' 1.1 1.2]
 ['' '' 'is_forecast' 0.0 0.0 0.0 0.0 0.0 0.0 0.0 1.0 1.0]
 ['3. УРАВНЕНИЯ' '' '' '' '' '' '' '' '' '' '' '']
 ['' '' 'x[t] = x_fut[t]' '' '' '' '' '' '' '' '' '']
 ['' '' 'y[t] = y[t-1] * y_rog[t]' '' '' '' '' '' '' '' '' '']]"""

def test_make(): 
   ar1 = get_resulting_workbook_array_for_make(abs_filepath)
   assert make_string == np.array_str(ar1)

def test_update():
   ar, equations_dict = get_array_and_support_variables(abs_filepath, sheet, pivot_col)         
   ar = fill_array_with_excel_formulas_based_on_is_forecast(ar, equations_dict, pivot_col)  
   assert update_string == np.array_str(ar)