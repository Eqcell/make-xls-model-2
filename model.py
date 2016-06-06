"""
Main functionality: 
- fill cells in Excel sheet with formulas (e.g. '=C3*D4') 
  based on list of variable names and equations. Formulas 
  go only to forecast periods columns where is_forecast == 1 

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

Todo later
----------
 - code review
 - make default sheet number/name a class variable + other changes related to sheet handling
 - test 'formulas' with py.test

"""

import pandas as pd
import numpy as np
from collections import OrderedDict
import re

from basefunc import to_rowcol, to_xl_ref

# from 'GDP[t-1]' catches 't-1'
T_ONLY_REGEX = r'\[([t+\-\d]+)\]'

# from '... + GDP[5] + 1' catches 'GDP[5] '
SEGMENT_REGEX = r'(\w+\[\d+\])'

# from 'GDP[5]' catches 'GDP', '5' 
VAR_PERIOD_REGEX = r'(\w+)\[(\d+)\]' 


class FormulaSegment(object):
    
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
        return to_xl_ref(self.row, self.col + self.column_offset, base=1)


class Formula(object):
    """
    Holds equation string and positioning information (var_to_rows, anchor) for
    dependent variable and allows to obtain corresponding Excel formula.
   
    Methods
    -------
    
    get_xl_formula(period)
    
       Example:
       xl_ref = Formula(equation_string, var_to_rows).get_xl_formula(period = 1)

    """
    
    def __init__(self, equation_string, var_to_rows, anchor="A1"):
        """
        Parameters
        ----------
        equation_string : equation for variable as text string
        var_to_rows : dictionary mapping variable names to rows on Excel sheet. Row numbers are based at 1. 
        anchor : A1-style reference to upper-left corner of the data block on Excel sheet, defaults to 'A1'
    
        """
        
        self.equation_string = self.strip_all_whitespace(equation_string)
        print(self.equation_string)
        self.equation_string = self.expand_shorthand(self.equation_string, var_to_rows)
        self.var_to_rows = var_to_rows 
        self.anchor = anchor
        
    def get_xl_formula(self, time_period):  
        
        indexed_equation = self.evaluate_time_indices(self.equation_string, time_period)         

        segments = re.findall(SEGMENT_REGEX, indexed_equation)   
        for seg_text in segments:
             xl_ref = FormulaSegment(seg_text, self.var_to_rows, self.anchor).xl_ref()
             rx = r'\b' + re.escape(seg_text)  # match beginning of word
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
                                '[' + str(period_offset) + ']')
        return text         


class Equations(object):
    
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
        varname = left_hand_side_expression.replace(" ", "").replace("[t]", "")
        return varname, formula                     

    @staticmethod    
    def error_duplicate_equation(key, eq1, eq2):
        raise ValueError("Two equations for the same variable. " + 
                         "\nVariable: " + key +          
                         "\nExisting equation: " + eq1 +
                         "\nAlternative equation: " + eq2)


class MathModel(object):
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
        assert 'is_forecast' in self.dataset.columns
        #    + check if enough data for equations were given
        #    + check if there are left-hand variables in equations without prior data 
        pass
        
    def _validate_positioning(self):     
        # Validating Excel positioning model:        
        #    verify if all required row locations were supplied
        pass

    def set_xl_positioning(self, var_to_rows, anchor="A1"):
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
                xl_dataset.loc[xl_dataset.index[i], varname] = Formula(
                    equation, self.var_to_rows, self.anchor).get_xl_formula(period_n)
                        
        return xl_dataset
                       
    
if __name__ == "__main__":    
    
    from basefunc import is_equal    
    
    # test data     
    COLUMNS = ['is_forecast', 'y', 'rog']   
    VAR_TO_ROWS = {'is_forecast': 2, 'y': 3, 'rog': 4}
    DF = pd.DataFrame({'y':           [85, 100, np.nan],
                       'rog':         [np.nan, np.nan,   1.05],
                       'is_forecast': [0, 0, 1]},
                      index=[2014, 2015, 2016])[COLUMNS]
    assert is_equal(DF, pd.read_excel('xl.xls').transpose()[COLUMNS])
    EQS = ['y = y[t-1] * rog'] 
    REF_DF = DF.copy()
    REF_DF.loc[2016, 'y'] = '=C3*D4'

    # test segment "GDP[1]" conversion to 'B5' 
    fs = FormulaSegment("GDP[1]", {'GDP': 5}, anchor="A1")
    assert fs.col == 1
    assert fs.row == 5
    assert fs.column_offset == 1
    assert fs.xl_ref() == 'B5'

    # formula strings converted to xl      
    pos = VAR_TO_ROWS, "A1"    
    # time period 1 + column offset 1 = B     
    assert "=B2" == Formula('is_forecast[t]', *pos).get_xl_formula(time_period=1)
    # time period 3 + column offset 1 = D
    assert '=C3*D4' == Formula('y[t-1] * rog', *pos).get_xl_formula(time_period=3) 
    # testing whitespace stripped
    assert "GDP[t]" == Formula("  GDP[t]  ", *pos).__repr__()
    
    # model with no Excel, local variables only
    m = MathModel(equations=EQS, dataset=DF)
    m.set_xl_positioning(var_to_rows=VAR_TO_ROWS)
    assert is_equal(m.get_xl_dataset(), REF_DF)
