# coding: utf-8
"""
   Generate Excel file with ordered rows containing Excel formulas 
   that allow to calculate forecast values based on historic data, 
   equations and forecast parameters. Order of rows in Excel file 
   controlled by template definition. Start year specified as input.

   Input:  
        data
        equations
        names
        controls (forecast parameters)
        formats 
           xl_filename
           sheet
           start_year
           row_labels        
        
        
   Output: 
        macro.xls
        (now - an array of values to be written to macro.xls)
"""

import numpy as np
import pandas as pd

from data_source import get_historic_data_as_dataframe, get_names_as_dict
from data_source import get_equations
from data_source import get_controls_as_dataframe
from data_source import get_row_labels, get_years_as_list, get_xl_filename

    
# when set to seros, the cellblock on output file will start at A1
CELLBLOCK_OFFSET_BY_ROW = 0 #5 
CELLBLOCK_OFFSET_BY_COL = 0 #2

def create_empty_cellblock_array(row_labels, years):   
   max_rows = 1 + len(row_labels) + CELLBLOCK_OFFSET_BY_ROW   
   max_col  = 2 + len(years) + CELLBLOCK_OFFSET_BY_COL 
   return np.empty((max_rows, max_col), dtype = object)

def populate_cellblock_with_labels_and_titles(cellblock, row_labels, names):   
    for j, label in enumerate(row_labels):
        # writing (label, names[label]['title']) to first two columns
        # better - write to temporary cellblock, not offsetted, 
        #          but make actual(pffsetted) cellblock before populate_cellblock_with_equations()        
        #          where actual offsets will be needed
        cellblock[1 + CELLBLOCK_OFFSET_BY_ROW + j, CELLBLOCK_OFFSET_BY_COL]     = names[label]['title']
        cellblock[1 + CELLBLOCK_OFFSET_BY_ROW + j, CELLBLOCK_OFFSET_BY_COL + 1] = label
    return cellblock
        
def populate_cellblock_with_years(cellblock, years):
    cellblock[0 + CELLBLOCK_OFFSET_BY_ROW,:] = [""] * CELLBLOCK_OFFSET_BY_COL + ["",""] + years
    # when not offsetted:
    # cellblock[0,:] = ["",""] + years
    return cellblock
    
def populate_cellblock_with_data(cellblock, row_labels, data):    
    # todo
    data = get_historic_data_as_dataframe()

    # Since we're dealing with a list of lists it's easier to just
    # use pure python loops
    years = cellblock[0][2:]
    props = [row[1] for row in cellblock[1:]]
    
    col_offset = 2
    row_offset = 1
    for i, y in enumerate(years):
        for j, p in enumerate(props):
            try:
                cellblock[j + row_offset][i + col_offset] = data[y][p]
            except KeyError:
                pass
    return cellblock
    
def populate_cellblock_with_controls(cellblock, row_labels, row_titles):
    # todo
    return cellblock
    
def populate_cellblock_with_equations(cellblock, row_labels, equations):
    return cellblock   
        
def dump_cellblock_to_xls(xl_filename, cellblock):
    print("\nThis will dump cellblock to " + xl_filename)  
    print(cellblock)

if __name__ == "__main__":      
    #init external variables
    data = get_historic_data_as_dataframe()  
    names = get_names_as_dict()
    equations = get_equations()
    controls = get_controls_as_dataframe()
    row_labels = get_row_labels()
    years = get_years_as_list()
      
    cellblock = create_empty_cellblock_array(row_labels, years)    
    cellblock = populate_cellblock_with_labels_and_titles(cellblock, row_labels, names)    
    cellblock = populate_cellblock_with_years(cellblock, years)
    
    #todo: make two functions below insert actual data 
    #      note - must use *years* and *row_labels* as pivot    
    cellblock = populate_cellblock_with_data(cellblock, row_labels, data)
    cellblock = populate_cellblock_with_controls(cellblock, row_labels, controls)   
    
    # --- Current output is: 
    # This will dump cellblock to macro.xls
    # [['' '' 2013 2014 2015 2016]
    # ['ВВП' 'GDP' None None None None]
    # ['Дефлятор ВВП' 'GDP_IP' None None None None]
    # ['Индекс физ.объема ВВП' 'GDP_IQ' None None None None]]
    
    # after changes are implemented only two None's will be there
    # ### indicates a digit
    # [['' '' 2013 2014 2015 2016]
    # ['ВВП' 'GDP' ### ### None None]
    # ['Дефлятор ВВП' 'GDP_IP' ### ### ### ###]
    # ['Индекс физ.объема ВВП' 'GDP_IQ' ### ### ### ###]]
    
    # not todo:
    cellblock = populate_cellblock_with_equations(cellblock, row_labels, equations)    
    
    xl_filename = get_xl_filename()
    dump_cellblock_to_xls(xl_filename, cellblock)
