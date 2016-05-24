#Implementation:
# - requirement - Windows machine with Microsoft Excel
# - Anaconda package suggested for libraries
# - Python 3.5 

#For improvement(not todo now):
# - make default sheet number/name a class variable
# - test 'formulas' with py.test

"""
DESCRIPTION
User story: 
  - the user wants to automate filling formulas on Excel sheet 
  - Excel sheet has simple 'roll forward' forecast - a spreadsheet model with some 
    historic variables and some control parameters for forecast (e.g. rates of growth) 
  - equations link control variables and previous period historic values to forecast values
  - equations are written down in excel sheet as text strings like 'y = y[t-1] * rog'
  - by running the python script the user has formulas filled in the Excel where necessary
  - the benefit is to have all model's formulas written down explicitly and not hidden in cells
  - currently we read input Excel sheet and write output sheet to different file or different sheet,
    but mya also write to same sheet to fill formulas
  
 
Some rules: 
  - from equations we know which variables are 'depenendent'('left-hand side')
  - control parameters are right-hand side variables, which do no appear on left side
  - all control variables must be supplied on sheet in dataset
  - we need explicit specification of year when the forecast starts -  by 'is_forecast' vector 
      
Simplifications/requirements:
  critical, but not checked: 
  - time series in rows only, horizontal orientation 
  - dataset starts at A1 cell
  checked:
  - must have 'is_forecast' vector in dataset
  not critical:
  - datablock is next to variable labels
  - time labels are years, not checked for continuity

Main functionality: fill cells in Excel sheet with formulas (e.g. '=C3*D4') based on 
                    list of variable names and equations.
                    
                    Formulas go only to forecast periods columns (where is_forecast == 1) 

Input Excel sheet:
-------------------------------
            A     B     C     D
1        year  2014  2015  2016
2 is_forecast     0     0     1
3           y    85   100   
4         rog              1.05
6 y = y[t-1] * rog
--------------------------------

Output Excel sheet:
-------------------------------
            A     B     C     D
1        year  2014  2015  2016
2 is_forecast     0     0     1
3           y    85   100  =C3*D4 
4         rog              1.05
6 y = y[t-1] * rog
--------------------------------

Comment:
- 'year' is not used in calculations 
- 'is_forecast' denotes forecast time periods, it is 0 for historic periods, 1 for forecasted
- 'y' is data variable
- 'rog' is control parameter
- 'y = y[t-1] * rog' is formula (equation)
"""

import pandas as pd
import numpy as np
from xlwings import Workbook, Range, Sheet
import os

from formulas import to_equations_dict, parse_equation_to_xl_formula

class Model():
    """    
    Fill dataset with formulas containing A1 cell references based on 
    equations and variable locations on Excel sheet: 
    
        dataset_with_formulas = Model(equations, dataset, var_to_rows).xl_dataset
    
    """
    
    def _validate_is_forecast(self):
        assert 'is_forecast' in self.dataset.columns
     
    def _validate_math_model(self):                
        # Validating mathematic model:
        #    check if enough data for equations were given
        #    check if there are left-hand variables in equations without prior data 
        pass
        
    def _validate_positioning(self):     
        # Validating Excel positioning model:        
        #    verify if all required row locations were supplied
        pass
    
    def __init__(self, dataset, equations, var_to_rows, varname_column = "B"):
        
        # warning: using dataset.copy() to prevent the global arguement from being modified inside the class
        self.dataset = dataset.copy()
        self.equations = to_equations_dict(equations) 
        self.var_to_rows = var_to_rows
        self.varname_col_n = col_to_num(varname_column) # column 2, not used
        
        self._validate_is_forecast()
        self._validate_math_model()
        self._validate_positioning()
        
        # main functionalty 
        self.xl_dataset = self.make_xl_dataset()
        
    def make_xl_dataset(self):
        
        xl_dataset = self.dataset 
        
        # rows where is_forecast == 1        
        forecast_index_positions = [t for t, flag in enumerate(xl_dataset.is_forecast) if flag == 1]
        
        # for each ariable name on left hand side of equations...          
        for varname in self.equations.keys():
            # know thire formula
            formula = self.equations[varname]
            # .. go over forecast time periods... 
            for i in forecast_index_positions:
               # later: must change 'period' to 'current_column_n' and review 'period' in 'parse_equation_to_xl_formula'
               period = i + 1
               # .... and assign formulas in xl_dataset
               xl_dataset.loc[xl_dataset.index[i], varname] = \
                        parse_equation_to_xl_formula(formula, self.var_to_rows, period)
                        
        return xl_dataset
                       
    def get_xl_dataset(self):         
        return self.xl_dataset       
        
def col_to_num(col_str):
    """ Convert base26 column string to number. """
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num
    
class SingleSheet():
    """
    Import inputs from Excel model sheet and write back formulas to it.
    Interface between class Model() and Excel file.
    
    Typical use:
    
    SingleSheet("xl.xls").save("xl_out.xls")
    
    """
    
    def __init__(self, filepath, sheet = 0):
        
        self._readsheet(filepath, sheet)        
        self.model = Model(self.dataset, self.equations, self.var_to_rows)        
        self._merge_formulas(self.model.xl_dataset)        
    
    def _readsheet(self, filepath, sheet):
        
        # all sheet content       
        self.image_df = pd.read_excel(filepath, sheet) 
        
        # dataset
        self.dataset = self.image_df.transpose()
        
        # equations
        self.equations = []        
        for label in self.dataset.columns:
            if "=" in label:
                self.equations.append(label)
                self.dataset = self.dataset.drop(label, 1)
            elif " " in label.strip():
                self.dataset = self.dataset.drop(label, 1)
               
        # var to rows        
        # +1 for rebasing from 0 and +1 fo header            
        self.var_to_rows = {l:i+1+1 for i, l in enumerate(self.image_df.index)}
        varnames = self.dataset.columns
        self.var_to_rows = {k:self.var_to_rows[k] for k in varnames}
  
    def _merge_formulas(self, xl_dataset):
        self.image_df = self.image_df.transpose()
        for col in xl_dataset.columns:
            self.image_df[col]=xl_dataset[col]
        self.image_df = self.image_df.transpose().fillna("")
    
    def save(self, filepath, sheet_n = 1):
        # later: filename may be not provided, must write to input sheet 
        
        def get_abspath(filepath):      
            folder = os.path.dirname(os.path.abspath(__file__))
            if not os.path.split(filepath)[0]:
                # provided filepath is file name only  
                return os.path.join(folder, filepath)
            else:
                # provided filepath is long path
                return filepath
            
        abspath = get_abspath(filepath)        
        wb = Workbook(abspath)
        # later: must check sheet_n exists
        Sheet(sheet_n).activate()
        # later: move 'A1' to CORNER_CELL = 'A1'
        Range(sheet_n, 'A1').value = to_matrix(self.image_df)
        wb.save()
        
        
def to_matrix(df): 
    df = df.fillna("")
    line0 = [""] +  df.columns.tolist()
    lines = [line0]
    for ix in df.index:
       def to_int(x):
           if type(x) == float and round(x) == x:
               return int(x)
           else:
               return x
       row = [to_int(x) for x in df.loc[ix,].tolist()] 
       lines.append([ix] + row) 
    return [[str(x) for x in line] for line in lines]    

        
if __name__ == "__main__":    
    
    def is_equal(df1, df2):
        # in numpy/pandas nan == nan is False, must substitute nans to compare frames
        # also 1 == 1.0 is false
        # below will only compare identically-labeled DataFrame objects, exceptions if different rows of columns
        flag = df1.fillna("") ==  df2.fillna("") 
        return flag.all().all()
        
    def make_test_data():
        df = pd.DataFrame({  'y' : [    85,    100, np.nan],
                           'rog' : [np.nan, np.nan,   1.05],
                   'is_forecast' : [     0,      0,      1]}, 
                           index = [  2014,   2015,   2016])  
        # guarantee column order                   
        df = df[['is_forecast', 'y', 'rog']]
        equations = ['y = y[t-1] * rog'] 
        var_to_rows = {'is_forecast': 2, 'y' : 3, 'rog' : 4}    

        ref = df.copy()
        ref.loc[2016,'y'] = '=C3*D4'
        
        return df, equations, var_to_rows, ref 
    
    # Testing with no Excel, local variables only
    df, equations_, var_to_rows_, ref_df = make_test_data()
    m = Model(equations = equations_, dataset = df, var_to_rows = var_to_rows_)
    model_df = m.get_xl_dataset()
    assert is_equal(model_df, ref_df)
    
    # Excel equals local variables
    xl_df = pd.read_excel('xl.xls').transpose()[['is_forecast', 'y', 'rog']]
    assert is_equal(df, xl_df) 
    
    # stand alone methods (must move them to existing classes?)    
    assert to_matrix(df) == [['', 'is_forecast', 'y', 'rog'], ['2014', '0', '85', ''], ['2015', '0', '100', ''], ['2016', '1', '', '1.05']]
    assert col_to_num("B") == 2

    # end-to-end call 
    SingleSheet('xl.xls').save('xl_out.xls')    
    SingleSheet('xl.xls').save('xl.xls', 2)
    # later: must check contents of output sheet to make end-to-end test complete    
    
    # Excel-based test
    ms = SingleSheet('xl.xls')
    assert ms.dataset.to_csv() == ',is_forecast,y,rog\n2014,0.0,85.0,\n2015,0.0,100.0,\n2016,1.0,,1.05\n'
    assert ms.var_to_rows == {'is_forecast': 2, 'y' : 3, 'rog' : 4}        
    assert ms.equations == ['y = y[t-1] * rog']
    assert ms.model.equations['y'] == 'y[t-1] * rog'
    assert is_equal(ref_df, ms.model.get_xl_dataset()) 
    
    #later: not sure it is 1, must switch to     
    assert '=B10' == parse_equation_to_xl_formula('credit', {'credit':10}, 1)
    #assert '=B2+A3*100' == parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100'
    #                       , {'GDP': 1, 'GDP_IQ': 2, 'GDP_IP': 3}, 1)
    