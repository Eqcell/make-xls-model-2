#!/usr/bin/env python3
"""
Creating xls test file for procedure that creates Excel sheet based on data, formulas and control parameters. 

Inputs:
  отчетные данные (data) - ряды данных с названиями переменных 
  формулы (formulas) -  выражения, которые определяют прогнозные значения следующего периода
  параметры (controls) - управляющие параметры, которые используются при расчете прогнозных значений c помощью формул, может задаваться как p_*
  
  not used:
  размещение (layout) - расположение рядов данных в итоговом лите Excel
  
Основная программа должна вызвать процедуру, которая создает лист 'result' по вводным данным, перечисленным выше. 
В данном файле такая процедура не вызывается, но создается файл, в котором показаны результаты работы процедуры.

Limitations:
  1. new variables cannot be created in formulas (left-hand side of formulas must have existing variable)
  2. will not writing variable text descriptions to 'result' sheet, only variable labels
"""

import pandas as pd

#
# Code has three parts: 
#     1. Creating and printing variables for 'data', 'formulas', 'controls' and 'result' sheets
#     2. Creating a test xls file with variables written to these sheets
#     3. Reading reference file to compare its content to variable values from (1)
#



#
#     1. Creating and printing variables for 'data', 'formulas', 'controls' and 'result' sheets
#

# --------------------------------------------------------------------
# Data
obs_years = [2010, 2011,  2012,  2013, 2014]
obs_x     = [100,   101, 102.5, 100.8, 95.5]
obs_y     = [3360, 2700,   500,  1200, 4800]
data = pd.DataFrame({'x': obs_x, 
                     'y': obs_y}, index = obs_years)
print("\nData:", data)

# --------------------------------------------------------------------
# Controls
forecast_years = [2015, 2016] #[obs_years[-1] + x for x in range(1,3)] 
y_rog = [1.1,  1.2]
x_fut = [102.5, 105]
controls = pd.DataFrame({'y_rog': y_rog,
                         'x_fut': x_fut}, index = forecast_years)
print("\nControls:", controls)

# --------------------------------------------------------------------
# Formulas
eq1 = "x[t] = x_fut[t]"
eq2 = "y[t] = y[t-1] * y_rog[t]"
# LIMITATION: variable z[t] must be in data before use, cannot add new variables by equationas of now
# eq3 = "z[t] = x[t] / y [t]" 
formulas = [eq1, eq2]
print("\nFormulas:", formulas)

# --------------------------------------------------------------------
# Resulting dataframe with values
all_years = obs_years + forecast_years
is_forecast = [0 for x in obs_years] + [1 for x in forecast_years]
data = data.reindex(index = all_years)
controls = controls.reindex(index = all_years)

#
#     2. Creating a test xls file with variables written to these sheets
#
iterator = zip([(t, isf) for t, isf in enumerate(is_forecast)], all_years)
for t, year in enumerate(all_years):
    if is_forecast[t]:
        # mimic formulas:
        data.x[year] = controls.x_fut[year]
        data.y[year] = data.y[year-1] * controls.y_rog[year]

output_layout = {'sheet': 'result',
                 'upper_left_corner': 'B2',
                 'variable_list': ['x', 'y', 'x_fut', 'y_rog']
                }

var_list = output_layout['variable_list'] 
out = data.join(controls)[var_list]
print("\nResulting dataframe with values:", out)

def write_sheet(df, sheet_name, writer):
    df.transpose().to_excel(writer, sheet_name)
def write_formulas(formulas_as_list, writer):
    pd.DataFrame(formulas_as_list).to_excel(writer, 'formulas', 
                 index = False)
def read_df(file, sheet_name):
    return pd.read_excel(file, sheet_name).transpose()
def read_formulas(file):
    df = pd.read_excel(file, 'formulas', index_col=None)
    return df[0].values.tolist()

# write to file 
writer = pd.ExcelWriter('testfile.xls')
write_sheet(data, 'data', writer)
write_sheet(controls, 'controls', writer)
write_formulas(formulas, writer)
write_sheet(out, 'result', writer)
writer.save()
# NOTE: 'testfile.xls' manually renamed 'ref_file.xls'

#
#     3. Reading reference file to compare its content to variable values from (1)
#
REF_FILE = "ref_file.xls"
assert read_df(REF_FILE, 'data').equals(data)
assert read_df(REF_FILE, 'controls').equals(controls)
assert read_df(REF_FILE, 'result').equals(out)
assert read_formulas(REF_FILE) == formulas

# TODO:
# - в папке src наладить запуск заполнения формул для 
# - наладить запуск python mxm.py ref_file.xls
# - предложить формальные тесты работы make-xls-model/src/make_xl_model.py
# - после утверждения реализовать тесты, желательно в py.test / unittest

# что не пока нравится:
# структура папки src
# в коде make_xl_model.py немного перемешаны уровни абстракции
# зависимость от xlwings
# короткий, но уже сложный интерфейс командной строки
# недостаточно примеров использования




