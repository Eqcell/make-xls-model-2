# coding: utf-8

import numpy as np
import pandas as pd
import re
from pprint import pprint

from data_source import print_specification 
from eqcell_core import parse_equation_to_xl_formula, TIME_INDEX_VARIABLES

###########################################################################
## Import test examples
###########################################################################
    
def get_target_ar():
    from data_source import _sample_for_xfill_array_after_equations
    return _sample_for_xfill_array_after_equations()
    
def get_sample_df_before_eq():
    from data_source import _sample_for_xfill_dataframe_before_equations
    return _sample_for_xfill_dataframe_before_equations()

def get_mock_specification():
    from data_source import get_mock_specification as _gmf
    return _gmf()    

###########################################################################
## Entry point + checks
###########################################################################
def make_wb_array2(data_df, controls_df, equations_list, var_label_list):
    # assemble inputs into dataframe
    df = get_dataframe_before_equations(data_df, controls_df, var_label_list)
     
    # get array with NaN at equation cells
    # WARNING: adds one column and one row, there are more hard-coded references to these rows/cols
    ar = get_array_before_equations(df)    
    
    # get resulting array 
    return fill_array_with_excel_formulas(ar, equations_list) 

def make_wb_array(model_spec, view_spec):
    """Creates an array, representing Excel worksheet based on *model_spec*, *view_spec*.
    Returns ndarray, which is to be dumped to Excel later in script."""
        
    # unpack variables locally
    # WARNING: names_dict not used
    [data_df, names_dict, equations_list, controls_df] = [s[1] for s in model_spec]
    [xl_file, sheet, var_label_list] = [s[1] for s in view_spec]
        
    # assemble inputs into dataframe
    df = get_dataframe_before_equations(data_df, controls_df, var_label_list)
     
    # get array with NaN at equation cells
    # WARNING: adds one column and one row, further code has hard-coded references to these rows/cols
    ar = get_array_before_equations(df)    
    
    # get resulting array 
    ar = fill_array_with_excel_formulas(ar, equations_list)     

    return ar

def check_wb_array():
    """
    >>> check_wb_array()
    True
    """
    m, v = get_mock_specification()
    ar = make_wb_array(m, v)
    target_ar = get_target_ar() 
    return np.array_equal(ar, target_ar)

    
###########################################################################
## get_dataframe_before_equations +  checks
###########################################################################
    
def check_get_dataframe_before_equations():
    """
    >>> check_get_dataframe_before_equations()
    True
    """
    model_spec, view_spec = get_mock_specification()    
    [data_df, names_dict, equations_list, controls_df] = [s[1] for s in model_spec]      
    df1 = get_sample_df_before_eq()
    df2 = get_dataframe_before_equations(data_df, controls_df, var_label_list)  
    return df1.equals(df2)

def get_dataframe_before_equations(data_df = None, controls_df = None, var_label_list = None):    
    """ Dataframe before equations obtained by merging historic data (*data_df*) and future values of control variables
        (*controls_df*), subsetted by *var_label_list*.
    """
    # Current behavior: 
    #       must merge data_df, controls_df into a common dataframe
    #       years are extended to include both data_df years and controls_df years
    #       missing values are None/Nan
    #       order of columns is same as listed in var_label_list
    # LATER:
    #       not todo: resove possible conflicts in data_df/control_df columns and  var_label_list   
    #       not todo: default behaviour in column first lists data_df, then elements of controls_df, which are not in controls_df ('is_forcast' in example)
    #       not todo: no check of years continuity
    
    # We first concatenate columns
    df = pd.concat([data_df, controls_df])
    
    # Subsetting a union of 'data_df' and 'controls_df', protected for error.
    try: 
       return df[var_label_list]
    except:
       print ("Error handling dataframes in get_dataframe_before_equations() in xl_fill.py")
       # LATER: add actual name of this file, obtained as a function
       return None
    
###########################################################################
## get_array_before_equations(df)
###########################################################################

def get_array_before_equations(df):
    """
       Decorate *df* with extra row (years) and column (var names) 
       and return as ndarray with object types. In resulting array some values for years
       will be NaN/nan. These are cells where Excel formulas need to be inserted.      
      
       Note: array of this kind directly represents an Excel worksheet.
             the intent is too fill this array and write it to Excel worksheet.
       
       Not todo: decorate also with a column of variable text descriptions (first column)
       """    
    ar = df.as_matrix().transpose().astype(object)
    labels = df.columns.tolist() 
    ar = np.insert(ar, 0, labels, axis = 1)
    years = [""] + df.index.astype(str).tolist()
    ar = np.insert(ar, 0, years, axis = 0)
    return ar

###########################################################################
## get_xl_formula - interface to eqcell_core.py and new text parser
###########################################################################

def build_formula(var_name, equations_list):
    """Returns a string like 'x(t-1) + 1' as a formula.""" 
    # DONE:          I consider splitting 'equations_list' into 'formulas_dict' a common task 
    #                for both parsers, but in concept this does not really belong to 'xl_fill.py'
    #                
    #                May later move strip_timeindex(str_), test_parse_to_formula_dict() and 
    #                parse_to_formula_dict(equations) to a separate module, callable either by two parsers OR
    #                at stage where we evaluate user-defined input. 
    import equations_preparser as eq_pp
    formulas_dict     = eq_pp.parse_to_formula_dict(equations_list)
    return eq_pp.get_formula(var_name, formulas_dict)



def get_xl_formula(cell, var_name, equations_list, variables_dict):
    """Returns a valid Excel formula as string, eg '=A2*C3'.   """ 
    # LATER:  - we currently do not check formulas: 
    #            * Only current period vars ьгые иу allowed on left side. Valid: "x(t)= x(t-1)". Not Valid: "x(t+1) = x(t)"
    #            * All required varables must be covered by formulas
    #            * What happens if there is a variable that is created from other variables, but not listed in data or controls?
    #          - cross-dependencies of 'TIME_INDEX_VARIABLES'       
    
    # get formula to work with
    formula_as_string =  build_formula(var_name, equations_list)
    # use sympy parser
    return get_xl_formula_sympyparser(cell, var_name, formula_as_string, variables_dict)

def get_xl_formula_textparser(cell, var_name, formula_as_string, variables_dict):
    # interface to new parsing algorithm 
    return None
        
def get_xl_formula_sympyparser(cell, var_name, formula_as_string, variables_dict):
    """
    cell is (row, col) tuple
    varname is like 'GDP'
    formula_as_string is 'x(t-1) + 1'
    variables_dict is like {'GDP': 2} - shows at which row the var is        
    """         
    try:
        time_period = cell[1]            
        return parse_equation_to_xl_formula(formula_as_string, variables_dict, time_period)
    except KeyError:
        return ""        

###########################################################################
## fill_array_with_excel_formulas
###########################################################################

def yield_cells_for_filling(ar):
    """
    Yields coordinates of nan values from data area in *ar* 
    Data area is all of ar, but not row 0 or col 0
          
    Example:
    
    >>> gen = yield_cells_for_filling([['', 2013, 2014, 2015, 2016],
    ...                                ['GDP', 66190.11992, 71406.3992, np.nan, np.nan],
    ...                                ['GDP_IP', 105.0467483, 107.1941886, 115.0, 113.0],
    ...                                ['GDP_IQ', 101.3407976, 100.6404858, 95.0, 102.5]])
    >>> next(gen)
    (1, 3)
    >>> next(gen)
    (1, 4)
    
    """
    row_offset = 1
    col_offset = 1
    
    # We loop and check which indexes correspond to nan
    for i, row in enumerate(ar[col_offset:]):
        for j, col in enumerate(row[row_offset:]):
            if np.isnan(col):
            # if math.isnan(col):
                yield i + col_offset, j + row_offset     

def get_variable_rows_as_dict(array, column = 0):
        variable_to_row_dict = {}        
        for i, label in enumerate(array[:,column]):           
            variable_to_row_dict[label] = i              
        #LATER: cut off one row (with years)
        #LATER: compare to full variable list
        return variable_to_row_dict
        # better - check is it is a valid variable name
        # var_list = unique(controls.columns.values.tolist() + row_labels)                

# def unique(list_):
    # """Returns unique elements from list.
    # >>> unique(['a','a'])
    # ['a']
    # """
    # return list(set(list_))     

        
def get_var_label(ar, row, var_column = 0):
        # WARNING: behavior not guaranteed, desired var_column may be not 0, must assign to constant
        return ar[row, var_column]
                
def fill_array_with_excel_formulas(ar, equations_list):        
        variables_dict = get_variable_rows_as_dict(ar)
        for cell in yield_cells_for_filling(ar):
            var_name = get_var_label(ar, cell[0])
            ar[cell] = get_xl_formula(cell, var_name, equations_list, variables_dict)
        return ar  
        
if __name__ == "__main__":   

    # unpack variables locally 
    from data_source import get_mock_specification
    model_spec, view_spec = get_mock_specification()
    [data_df, names_dict, equations_list, controls_df] = [s[1] for s in model_spec]
    [xl_file, sheet, var_label_list] = [s[1] for s in view_spec]
    
    # task formulation - inputs
    print("\n****** Module inputs:")    
    print_specification(model_spec)
    print_specification(view_spec)
    
    # task formulation - final result
    print("\n****** Task - produce array with values and string-like formulas (intent - later write it to Excel)")
    print("Target array:")
    
    target_ar = get_target_ar()
    
    print(target_ar) 
    # end task formulation   
    
    # Solution:
    print("\n******* Solution flow:")    
    print("*** Array before equations:")
    df = get_dataframe_before_equations(data_df, controls_df, var_label_list)
    ar = get_array_before_equations(df)
    print(ar)
    
    print("\n***  Split formulas:")
    
    import equations_preparser as eq_pp
    formulas_dict = eq_pp.parse_to_formula_dict(equations_list)
    pprint(formulas_dict)
    
    print("\n***  Assign variables to array rows:")
    variable_to_row_dict = get_variable_rows_as_dict(ar)
    pprint(variable_to_row_dict)
    
    print("\n***  Iterate over NaN in data area + fill with formulas:")    
    ar = fill_array_with_excel_formulas(ar, equations_list)    
    print("\n Resulting array:")
    print(ar)
         
    print("\n*** Must be like:")    
    print(target_ar)    
    is_equal = np.array_equal(ar, target_ar)         
    
    print("\n*** Solution complete: " + str(is_equal))
    print()
    print(np.equal(ar, target_ar))
    
    import doctest
    doctest.testmod()
