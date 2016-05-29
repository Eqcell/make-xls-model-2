# -*- coding: utf-8 -*-
"""
Created on Sun May 29 09:12:29 2016

@author: Евгений
"""

import numpy as np
import pandas as pd

import xlrd
from xlwings import Workbook, Range, Sheet
import os

from basefunc import to_rowcol
from model import MathModel

class ModelOnSheet():
    """
    Import inputs from Excel model sheet and write back formulas to it.
    Interface between class Model() and Excel file.
    
    Typical use:
    
    Sheet("xl.xls").save("xl_out.xls")
    
    """
    
    def __init__(self, filepath, sheet = 0, anchor = 'A1'):
        
        self._readsheet(filepath, sheet)        
        self.model = MathModel(self.dataset, self.equations).set_xl_positioning(self.var_to_rows)        
        self._merge_formulas(self.model.get_xl_dataset())        
    
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
        # +1 for rebasing from 0 and +1 for header            
        self.var_to_rows = {l:i+1+1 for i, l in enumerate(self.image_df.index)}
        self.var_to_rows = {k:self.var_to_rows[k] for k in self.dataset.columns}
  
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
        Range(sheet_n, 'A1').value = self.to_matrix(self.image_df)
        wb.save()
        
    @staticmethod    
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

    
def read_sheet_as_array(filename, n=0):
    """Converts n-th sheet from an Excel file into an ndarray"""
    contentstring = open(filename, 'rb').read()
    book  = xlrd.open_workbook(file_contents=contentstring)
    sheet = book.sheets()[n]
    array = np.empty((sheet.ncols, sheet.nrows), dtype=object)
    
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            value = sheet.cell(row, col).value
            # force type to 'int' where possible
            if isinstance(value, float) and round(value) == value:
                value = int(value)                
            array[row][col] = value
    
    return array
    
def read_range_as_df(filename, n=0, anchor="A1"):

    arr = read_sheet_as_array("xl2.xls")
    r, c = to_rowcol(anchor, base = 0)
    data = arr[r:,c:]
    
    return pd.DataFrame(data=data[1:,1:],    # values
                    index=data[1:, 0],    # 1st column as index
                  columns=data[0 ,1:])  # 1st row as the column names

def merge_df_in_array(arr, df, anchor):
    r, c = to_rowcol(anchor, base = 0)
    arr[r+1:,c+1:] = df.as_matrix()
    return arr 

if __name__ == "__main__":
    
    from basefunc import is_equal   
    
    arr = read_sheet_as_array("xl2.xls", n=0)
    df = read_range_as_df("xl2.xls", n=0, anchor="D3")
    df.iloc[0:,0:]= -1
    arr2 =  merge_df_in_array(arr, df, "D3")

    # Test data     
    COLUMNS = ['is_forecast', 'y', 'rog']   
    VAR_TO_ROWS = {'is_forecast': 2, 'y' : 3, 'rog' : 4}
    DF =  pd.DataFrame({  'y' : [    85,    100, np.nan],
                        'rog' : [np.nan, np.nan,   1.05],
                'is_forecast' : [     0,      0,      1]},
                        index = [  2014,   2015,   2016])[COLUMNS]   
    assert is_equal(DF, pd.read_excel('xl.xls').transpose()[COLUMNS])
    EQS = ['y = y[t-1] * rog'] 
    REF_DF = DF.copy()
    REF_DF.loc[2016,'y'] = '=C3*D4'

    # standalone method
    assert ModelOnSheet.to_matrix(DF) == [['', 'is_forecast', 'y', 'rog'], ['2014', '0', '85', ''], ['2015', '0', '100', ''], ['2016', '1', '', '1.05']]
    
    mos = ModelOnSheet('xl.xls')
    assert is_equal(mos.dataset, DF)
    assert mos.var_to_rows == VAR_TO_ROWS
    assert mos.equations == EQS
    assert mos.model.equations['y'] == 'y[t-1] * rog'
    assert is_equal(REF_DF, mos.model.get_xl_dataset()) 
    
    # end-to-end call 
    ModelOnSheet('xl.xls').save('xl_out.xls')    
    ModelOnSheet('xl.xls').save('xl.xls', 2)
    # later: must check contents of output sheet to make end-to-end test complete    


    