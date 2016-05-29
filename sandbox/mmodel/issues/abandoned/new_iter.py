import numpy as np

def yield_cell_coords_for_filling(ar, pivot_labels, pivot_col = 0):
    """
    Must yield coordinates of cells to the right of pivot_col, where
    pivot_col values is in pivot_labels, and cell value is NaN.
    
    Example:
    
    >>> gen = yield_cell_coords_for_filling([['', 2013, 2014, 2015, 2016],
    ...                                      ['GDP', 66190.11992, 71406.3992, np.nan, np.nan],
    ...                                      ['GDP_IP', np.nan, 107.1941886, 115.0, 113.0],
    ...                                      ['GDP_IQ', np.nan, 100.6404858, 95.0, 102.5]], 
    ...                                      ["GDP"])
    >>> next(gen)
    (1, 3, 'GDP')
    >>> next(gen)
    (1, 4, 'GDP')
    """
    for i, row in enumerate(ar):
        var_label = row[pivot_col]
        if var_label in pivot_labels:
            for j, ce in enumerate(row):
                if j > pivot_col:
                    if np.isnan(ce):
                        yield i, j, var_label

def get_variable_rows_as_dict(array, full_var_list, pivot_col = 0):
    variable_to_row_dict = {}        
    for i, label in enumerate(array[:,pivot_col]):
        if label in full_var_list:
            variable_to_row_dict[label] = i              
    return variable_to_row_dict

from xl_fill import get_xl_formula

def fill_array_with_excel_formulas(ar, equations_list, pivot_labels, all_labels):    
        variables_dict = get_variable_rows_as_dict(ar, all_labels)
        for i, j, var_label in yield_cell_coords_for_filling(ar, pivot_labels):
            ar[i, j] = get_xl_formula((i,j), var_label, equations_list, variables_dict)
        return ar 
        
if __name__ == "__main__":
    import doctest
    doctest.testmod()
