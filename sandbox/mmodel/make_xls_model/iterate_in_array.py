import numpy as np
from formula_parser import parse_equation_to_xl_formula

###########################################################################
## Iteration for --make
###########################################################################

def yield_cell_coords_for_filling(ar, pivot_labels, pivot_col):
    """Must yield coordinates of cells to the right of pivot_col, where
    pivot_col values is in pivot_labels, and cell value is NaN.
    
    Example:
    
    >>> gen = yield_cell_coords_for_filling([['', 2013, 2014, 2015, 2016],
    ...                                      ['GDP', 66190, 71406, np.nan, np.nan],
    ...                                      ['GDP_IP', np.nan, 107.1, 115.0, 113.0],
    ...                                      ['GDP_IQ', np.nan, 100.6,  95.0, 102.5]], 
    ...                                      ["GDP"], 0)
    >>> next(gen)
    (1, 3, 'GDP')
    >>> next(gen)
    (1, 4, 'GDP')
    """
    for i, row in enumerate(ar):
        var_label = row[pivot_col]
        if var_label in pivot_labels:
            for j, cell_content in enumerate(row):
                if j > pivot_col:
                    if np.isnan(cell_content):
                        yield i, j, var_label

def get_variable_rows_as_dict(array, pivot_col = 0):
#    """
#     import numpy as np   
#     var_dict = variable_to_row_as_dict(np.array([
#    ...       ['', '2013', '2014', '2015', '2016']
#    ...      ,['GDP', 66190, 71406, '=C2*D3/100*D4/100', '=D2*E3/100*E4/100']
#    ...      ,['GDP_IQ', 101.3407,  100.6404, 95.0,  102.5]
#    ...      ,['GDP_IP', 105.0467,  107.1941, 115.0, 113.0] ]
#    ...      , dtype=object))
#     var_dict['GDP']
#    1   
#    """
    variable_to_row_dict = {}        
    for i, label in enumerate(array[:,pivot_col]):
        if not "=" in label and not label.strip().startswith("#"):     
            variable_to_row_dict[label] = i
    return variable_to_row_dict

def fill_array_with_excel_formulas(ar, equations_dict, pivot_col = 0):    
    pivot_labels = equations_dict.keys()
    var_dict = get_variable_rows_as_dict(ar, pivot_col)    
    for i, j, var_name in yield_cell_coords_for_filling(ar, pivot_labels, pivot_col):
        formula_as_string = equations_dict[var_name]
        time_period = j 
        ar[i, j] = parse_equation_to_xl_formula(formula_as_string, 
                                                var_dict, time_period)
    return ar

###########################################################################
## Iteration for --update
###########################################################################

def columns_list_if_cellvalue_is_one(ar, row):    
    return [i for i, val in enumerate(ar[row,:]) if val == 1]

def yield_cell_coords_for_filling_based_on_is_forecast(ar, pivot_labels, pivot_col):
    var_dict = get_variable_rows_as_dict(ar, pivot_col)    
    for lab in pivot_labels:
        row = var_dict[lab]
        for col in columns_list_if_cellvalue_is_one(ar, var_dict['is_forecast']):
            yield row, col, lab

def fill_array_with_excel_formulas_based_on_is_forecast(ar, equations_dict, pivot_col):
    pivot_labels = equations_dict.keys()
    var_dict = get_variable_rows_as_dict(ar, pivot_col)      
    #print(equations_dict)
    #print(var_dict)
    for i, j, var_name in yield_cell_coords_for_filling_based_on_is_forecast(ar, 
                                                              pivot_labels, pivot_col): 
        formula_as_string = equations_dict[var_name]       
        time_period = j 
        ar[i, j] = parse_equation_to_xl_formula(formula_as_string, 
                                                var_dict, time_period)
    return ar

###########################################################################

if __name__ == "__main__":
    import doctest
    doctest.testmod()