"""

    Fill cells in Excel sheet with formulas (e.g. '=C3*D4') 
    based on list of variable names and equations as text strings.
    Formulas go only to forecast periods columns where is_forecast == 1. 

    
```
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
```

Comment:
- 'year' is not used in calculations 
- 'is_forecast' denotes forecast time periods, it is 0 for historic periods, 1 for forecasted
- 'y' is data variable
- 'rog' is control parameter
- 'y = y[t-1] * rog' is formula (equation)

"""


import pandas as pd
import numpy as np
from collections import OrderedDict
import re
import xlrd
from xlwings import Workbook, Range, Sheet
import argparse
import os


#----------------------------------------------------------------------------------
#
#    Excel cell reference functions
#
#----------------------------------------------------------------------------------

def to_xl_ref(row, col, base = 1):
    if base == 1:
        return xlrd.colname(col-1) + str(row)
    elif base == 0:
        return xlrd.colname(col) + str(row+1)

def col_to_num(col_str):
    """ Convert base26 column string to number. """
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num
    
def to_rowcol(xl_ref, base = 1):
    xl_ref = xl_ref.upper()
    letters, b =  re.search(r'(\D+)(\d+)', xl_ref).groups()        
    return int(b) + (base-1), col_to_num(letters) + (base-1) 
    
def is_equal(df1, df2):
    # in numpy/pandas nan == nan is False, must substitute nans to compare frames
    # also 1 == 1.0 is false
    # below will only compare identically-labeled DataFrame objects, exceptions if different rows of columns
    flag = df1.fillna("") ==  df2.fillna("") 
    return flag.all().all()

#----------------------------------------------------------------------------------
#
#    MathModel class
#
#----------------------------------------------------------------------------------
    
# from 'GDP[t-1]' catches 't-1'
T_ONLY_REGEX = r'\[([t+\-\d]+)\]'

# from '... + GDP[5] + 1' catches 'GDP[5] '
SEGMENT_REGEX = r'(\w+\[\d+\])'

# from 'GDP[5]' catches 'GDP', '5' 
VAR_PERIOD_REGEX = r'(\w+)\[(\d+)\]' 


class FormulaSegment():
    
    def __init__(self, seg_text, var_to_rows, anchor): 
        """
        Parameters
        ----------
        seg_text : string containing variable name and integer index in brackets, e.g. 'GDP[1]' 
        var_to_rows : dictionary mapping variable names to rows on Excel sheet. Row numbers are based at 1. 
        anchor : A1-style reference to upper-left corner of the data block on Excel sheet, defaults to 'A1'
        
        """
        
        varname, b = re.search(VAR_PERIOD_REGEX, seg_text).groups()
        self.col = int(b)
        if varname in var_to_rows.keys():
           self.row = var_to_rows[varname]
        else:
           raise KeyError("Variable without row: " + varname) 
           
        # anchor is used to calculate column offset
        # for "A1" offset is 1, which means first time period will be in column 2 (or "B")
        r, c = to_rowcol(anchor)
        self.column_offset = int(c)
    
    def xl_ref(self):
        """Returns A1-style reference for segment, eg. 'B5', 'D20', etc. """          
        return to_xl_ref(self.row, self.col + self.column_offset, base = 1)    
            
class Formula():
    """
    Holds equation string and positioning information (var_to_rows, anchor) for
    dependent variable and allows to obtain corresponding Excel formula.
   
    Methods
    -------
    
    get_xl_formula(period)
    
       Example:
       xl_ref = Formula(equation_string, var_to_rows).get_xl_formula(period = 1)

    """
    
    def __init__(self, equation_string, var_to_rows, anchor = "A1"):
        """
        Parameters
        ----------
        equation_string : equation for variable as text string
        var_to_rows : dictionary mapping variable names to rows on Excel sheet. Row numbers are based at 1. 
        anchor : A1-style reference to upper-left corner of the data block on Excel sheet, defaults to 'A1'
    
        """
        
        self.equation_string = self.strip_all_whitespace(equation_string)
        self.equation_string = self.expand_shorthand(self.equation_string, var_to_rows)
        self.var_to_rows = var_to_rows 
        self.anchor = anchor
        
    def get_xl_formula(self, time_period):  
        
        indexed_equation = self.evaluate_time_indices(self.equation_string, time_period)         

        segments = re.findall(SEGMENT_REGEX, indexed_equation)   
        for seg_text in segments:
             xl_ref = FormulaSegment(seg_text, self.var_to_rows, self.anchor).xl_ref()
             rx = r'\b' + re.escape(seg_text) # match beginning of word  
             indexed_equation = re.sub(rx, xl_ref, indexed_equation)
        return '=' + indexed_equation
        
    def __repr__(self):
        return self.equation_string
            
    @staticmethod
    def strip_all_whitespace(equation_string):
        return re.sub(r'\s+', '', equation_string) 
      
    @staticmethod    
    def expand_shorthand(text, var_to_rows):
        for var in [v for v in var_to_rows.keys() if v]:
           # catches GDP without further [] 
           rx = var + r'(?!\s*[\dA-Za-z_^\[])'
           text = re.sub(rx, var + '[t]', text)
        return text  

    @staticmethod    
    def evaluate_time_indices(text, time_period):
        
        for time_index_expression in re.findall(T_ONLY_REGEX, text):
            try:
                # 't' will be used inside time_index_expression                 
                t = time_period
                period_offset = eval(time_index_expression)
            except:
                raise ValueError('Time idex expression invalid: ' + time_index_expression)
            
            text = text.replace('[' + time_index_expression + ']', 
                                '[' + str(period_offset)    + ']')
        return text         

class Equations():
    
    def __init__(self, equation_strings):
        
        eq_dict = OrderedDict()
        # disregard comments and strings without '='
        equation_strings = [eq for eq in equation_strings 
                            if "=" in eq and not eq.strip().startswith("#")]
        for eq in equation_strings:
            key, formula = self.parse_equation_string(eq)
            if key in eq_dict.keys():
                self.error_duplicate_equation(key, eq_dict[key], formula)                    
            else:
                eq_dict[key] = formula.strip()
        self.dict = eq_dict  
        
    @staticmethod    
    def parse_equation_string(string):
        left_hand_side_expression, formula = string.split('=')
        varname = left_hand_side_expression.replace(" ","").replace("[t]", "")    
        return varname, formula                     

    @staticmethod    
    def error_duplicate_equation(key, eq1, eq2):
        raise ValueError("Two equations for the same variable. " + 
                         "\nVariable: " + key +          
                         "\nExisting equation: " + eq1 +
                         "\nAlternative equation: " + eq2)

class MathModel():
    """    
    Fill dataframe with formulas containing A1 cell references based on 
    equations and variable locations on Excel sheet. 
    
    Methods
    -------
    set_xl_postioning(var_to_rows, anchor)
    get_xl_dataset()
    
    """
    
    def __init__(self, dataset, equations):
       """
       Parameters
       ----------
       dataset : dataframe with time series for variables by year
       equations : list of text strings holding equations for variables
       
       """
        
       # warning: using dataset.copy() to prevent the global variable from being modified inside the class
       self.dataset = dataset.copy()
       self.equations = Equations(equations).dict 
       self._validate_math_model()
    
    def _validate_math_model(self):                
        # Validating mathematic model:
        #    + check if enough data for equations were given
        #    + check if there are left-hand variables in equations without prior data 
        pass
        
    def _validate_positioning(self):     
        # Validating Excel positioning model:        
        #    verify if all required row locations were supplied
        pass

    def set_xl_positioning(self, var_to_rows, anchor = "A1"):
        """
        Provide positoiing information about variable locations in rows ('var_to_rows') 
        and overall dataframe range location ('anchor').         
        
        Parameters
        ----------
        var_to_rows : dictionary mapping variable names to rows on Excel sheet. Row numbers are based at 1. 
        anchor : A1-style reference to upper-left corner of the data block on Excel sheet, defaults to 'A1'
    
        """

        self.var_to_rows = var_to_rows
        self.anchor = anchor
        self._validate_positioning()
        return self 
        
    def get_xl_dataset(self):
        
        xl_dataset = self.dataset 
        
        # rows where is_forecast == 1        
        forecast_index_positions = [t for t, flag in enumerate(xl_dataset.is_forecast) if flag == 1]
        
        # for each variable name on left hand side of equations...          
        for varname in self.equations.keys():
            # ... get formula for the variable ...
            equation = self.equations[varname]
            # .. go over forecast time periods... 
            for i in forecast_index_positions:
               # .... and assign formulas in xl_dataset
               period_n = i + 1
               xl_dataset.loc[xl_dataset.index[i], varname] = \
                        Formula(equation, 
                                self.var_to_rows,
                                self.anchor).get_xl_formula(period_n)
                        
        return xl_dataset
        

#----------------------------------------------------------------------------------
#
#    ExcelSheet class
#
#----------------------------------------------------------------------------------

       
def _get_xlrd_sheet(filename, sheet):
   
   contentstring = open(filename, 'rb').read()
   book = xlrd.open_workbook(file_contents=contentstring)
   
   if isinstance(sheet, int):
       # if 'sheet' is integer, we assume 'sheet' is based at 1   
       return book.sheet_by_index(sheet-1)
   elif isinstance(sheet, str) and sheet in book.sheet_names():
       return book.sheet_by_name(sheet)
   else:
       raise Exception("Cannot find sheet :" + str(sheet))
       
def get_array_from_sheet(filename, sheet):
    sheet = _get_xlrd_sheet(filename, sheet)       
    array = np.empty((sheet.nrows,sheet.ncols), dtype=object)
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            value = sheet.cell(row, col).value
            # force values type to 'int' where possible
            if isinstance(value, float) and round(value) == value:
                value = int(value)                
            array[row][col] = value
    return array              

def _fullpath(path):
    
    # current directory
    cur_dir = os.path.dirname(os.path.abspath(__file__))

    if os.path.normcase(os.path.normpath(path)).startswith(cur_dir):
       return path
    else:
       return os.path.join(cur_dir, path)

def write_array_to_sheet(filepath, sheet, arr):

    path = _fullpath(filepath) # Workbook(path) seems to fail unless full path is provided
    if os.path.exists(path):
        wb = Workbook(path)
        Sheet(sheet).activate()
        Range("A1").value = arr 
        wb.save()
    else:
        raise FileNotFound(path) 
   

class ExcelSheet():

    """

    Access Excel file for reading sheet and saving sheet with formulas.
    
    Notes
    -----
    - Operates on numpy array *self.arr* representing cells in Excel sheet. 
    - Uses MathModel class to populate formulas.   
    
    Methods
    -------
    .insert_formulas() method populates cells in forecast periods with Excel-style formulas. 
    .save() will read first sheet of Excel file and populate it with formulas.

    """
    
    def __init__(self, filepath, sheet = 1, anchor = 'A1'):
        """
        Inputs
        ------
        filepath : valid path to Excel file, xls only, xlsx not supported
        sheet: string or integer >=1, representing sheet name or number starting at 1, defaults to first sheet 
        anchor : string with A1 style reference, defaults to "A1"
        """ 
        
        self.source = {'path':filepath, 'sheet':sheet, 'anchor':anchor}
        self.arr = get_array_from_sheet(filepath, sheet)
        self.anchor_rowx, self.anchor_colx = to_rowcol(anchor, base = 0)

        self.dataset = self.extract_dataframe(self.arr, self.anchor_rowx, self.anchor_colx).transpose()
        self.check_dataset()
        self.equations = self.pop_equations()
        self.var_to_rows = self.get_variable_locations_by_row()
        self.model = MathModel(self.dataset, self.equations).set_xl_positioning(self.var_to_rows, anchor) 
        self.insert_formulas()
    
    def check_dataset(self):
        if not 'is_forecast' in self.dataset.columns:
             print("Datset columns:\n", self.dataset.columns) 
             print("\nAnchor cell row and column:\n", self.anchor_rowx, self.anchor_colx) 
             raise ValueError("Row 'is_forecast' not found in dataframe.\nPossible reason - wrong anchor cell.")     
                     
    @staticmethod
    def extract_dataframe(arr, anchor_rowx, anchor_colx):
        """Return a part of 'self.arr' starting anchor cell as dataframe.""" 
           
        data = arr[anchor_rowx:,anchor_colx:]
        #
        
        return pd.DataFrame(data=data[1:,1:],    # values
                           index=data[1:, 0],    # 1st column as index
                         columns=data[0 ,1:])    # 1st row as the column names

    def get_variable_locations_by_row(self):
        """Return dictionary with variable row locations.""" 
        var_to_rows = {}
        column_with_labels = self.arr[:,self.anchor_colx]
        for rowx, label in enumerate(column_with_labels):
            if label in self.dataset.columns:
                # +1 to rebase from 0  
                var_to_rows[label] = rowx + 1        
        return var_to_rows  
        
    def pop_equations(self):       
        """Return list of strings containing equations. 
           Also cleans self.dataset off junk non-variable columns""" 
        equations = []        
        
        def drop(label):
            if label in self.dataset.columns:
                self.dataset = self.dataset.drop(label, 1)
                
        for label in self.dataset.columns:
            if "=" in label:
                equations.append(label)
                drop(label)
            elif (" " in label.strip() 
                  or label.startswith("#")
                  or len(label) == 0):
                drop(label)
        return equations       

    def insert_formulas(self):
        """Populate formulas on array representing Excel sheet."""        
        df = self.model.get_xl_dataset()
        column_with_labels = self.arr[:,self.anchor_colx]
        for rowx, label in enumerate(column_with_labels):
            if label in df.columns:                
                self.arr[rowx,self.anchor_colx+1:] = df[label].as_matrix()
        return self

    def save(self, filepath=None, sheet=None):
        if not filepath:
            filepath = self.source['path']            
        if not sheet:
            sheet = self.source['sheet']
        self.target = {'path':filepath, 'sheet':sheet}
 
        write_array_to_sheet(filepath, sheet,  self.arr)
        return self
        
    def echo(self):
        print("\n  File: " + self.target['path'])
        print(  " Sheet: " + self.target['sheet']) 
        print("\nUpdated formulas:")
        eqs = ["    " + k + " = " + v  for k, v in self.model.equations.items()]
        for e in eqs:
            print(e)
        return self
        
def cli():
    """Command line interface to XlSheet(filepath, sheet, anchor).save()"""
    
    parser = argparse.ArgumentParser(description='Command line interface to XlSheet(filename, sheet, anchor).save()',
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('filename',                        help='filename or path to .xls file')
    parser.add_argument('sheet',  nargs='?', default=1,  help='sheet name or sheet index starting at 1')
    parser.add_argument('anchor', nargs='?', default='A1', help='reference to upper-left corner of data block, defaults to A1')
    
    args = parser.parse_args()

    filename = args.filename
    anchor = args.anchor

    try:
        sheet = int(args.sheet)
    except ValueError:
        sheet = args.sheet
   
    xl = ExcelSheet(filename, sheet, anchor).save()
    xl.echo()

if __name__ == "__main__":
    cli()       