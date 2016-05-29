from xlwings import Workbook, Range, Sheet
import numpy as np

    # workfile = "D:/make-xls-model-master/spec.xls"
    # sheet = "model"
    # wb = Workbook(workfile)
    # Sheet(sheet).activate()
    # ar = _sample_for_xfill_array_after_equations()
    # Range(sheet, 'A1').value = ar
    # wb.save()
    # wb.close()
    
from data_source import _sample_for_xfill_array_after_equations 
ar =  _sample_for_xfill_array_after_equations



def get_workbook(workfile, sheet):
    wb = Workbook(workfile)
    
    # work on given sheet
    Sheet('model').activate()
    
    # >>> import numpy as np
    # >>> wb = Workbook()
    # >>> Range('A1').value = np.eye(5)
    # >>> Range('A1', asarray=True).table.value
    # array([[ 1.,  0.,  0.,  0.,  0.],
           # [ 0.,  1.,  0.,  0.,  0.],
           # [ 0.,  0.,  1.,  0.,  0.],
           # [ 0.,  0.,  0.,  1.,  0.],
           # [ 0.,  0.,  0.,  0.,  1.]])
    return wb
        
def save_workbook(workbook, savepath=None):
    """
    Saves the workbook in given path or overwrites existing file.
    """
    if savepath is None:
        workbook.save()                          # save over the same workbook (overwrite)
    else:
        savepath = os.path.normcase(savepath)    # makes '/' into '\' for windows compatibility
        workbook.save(savepath)                  # SaveAs with given path

def close_workbook(workbook):
    """
    Closes the workbook in given path
    """
    workbook.close()
    


def write_array_to_xlsx_using_xlwt(ar, xlsx_path, sheet_name):
    ar = _sample_for_xfill_array_after_equations()
    wb = Workbook(workfile)
    Sheet(sheet_name).activate()
    Range(sheet_name, 'A1').value = ar

   

def apply_formulas_on_sheet(workbook, variables, parsed_formulas, start_cell):
    """
    Takes each cell in the sheet inside the rectangle formed by Start_cell and End_cell
    checks 1) if the cell is in a row with a variable as first element
           2) if the cell is in a column with `is_forecast=1`
    If all above conditions are met, then apply a fitting formula as obtained from find_formulas()
    Apply's the solution on the workbook cells. Raises error if any problem arises.
    input
    -----
    workbook:   Workbook xlwings object
    variables: A dict of variables from excel sheet
    parsed_formulas: A dict of formulas with key as row_index and value as dict of left-side and right-side sympy expressions
    start_cell: Start cell dictionary
    """
    workbook.set_current()    # sets the workbook as the current working workbook
    forecast_row = Range(get_sheet(), (variables['is_forecast'], start_cell['col'] + 1)).horizontal.value
    col_indices = [start_cell['col'] + 1 + index for index, el in enumerate(forecast_row) if el == 1]    # checks if is_forecast value in this col is = 1 and notes down col index
    row_indices = list(variables.values())
    row_indices.remove(variables['is_forecast'])

    for col, row in itertools.product(col_indices, row_indices):

        formula_dict = _get_formula(parsed_formulas, row, col)

        if formula_dict:
            dependent_variable_with_time_index = formula_dict['dependent_var']     # get expression for dependent variable, e.g. a(t)
            # dependent_variable_locations - values like {b(t): (8, 5)}
            dependent_variable_locations = simplify_expression(dependent_variable_with_time_index, col, variables)
            dv_key, dv_coords = dependent_variable_locations.popitem()

            # 2015-05-12 03:09 PM
            # --- Need to make this check elsewhere
            if dependent_variable_locations:
                raise ValueError('cannot have more than one dependent variable on left side of equation')
            # --- end

            # find excel type formula string
            right_side_expression = formula_dict['formula']
            formula_str = get_excel_formula_as_string(right_side_expression, col, variables)
            Range(get_sheet(), dv_coords).formula = formula_str                # Apply formula on excel cell

workfile = "D:/make-xls-model-master/spec.xls"
sheet = "model"
wb = Workbook(workfile)
Sheet(sheet).activate()
wb = get_workbook(workfile, sheet)
ar = _sample_for_xfill_array_after_equations()
Range(sheet, 'A1').value = ar
save_workbook(wb)	