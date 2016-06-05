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

class SheetImage():
    
    def __init__(self, arr, anchor):
    
        self.arr = arr
        self.anchor_rowx, self.anchor_colx = to_rowcol(anchor, base = 0)

        self.dataset = self.extract_dataframe().transpose()
        self.equations = self.pop_equations()
        self.var_to_rows = self.get_variable_locations_by_row(varlist=self.dataset.columns)
        self.model = MathModel(self.dataset, self.equations)\
                     .set_xl_positioning(self.var_to_rows, anchor)        

    def extract_dataframe(self):
        """Return a part of 'self.arr' as dataframe.""" 
           
        data = self.arr[self.anchor_rowx:,self.anchor_colx:]
        
        return pd.DataFrame(data=data[1:,1:],    # values
                           index=data[1:, 0],    # 1st column as index
                         columns=data[0 ,1:])    # 1st row as the column names

    def insert_formulas(self):
        df = self.model.get_xl_dataset()
        column_with_labels = self.arr[:,self.anchor_colx]
        for rowx, label in enumerate(column_with_labels):
            if label in df.columns:
                self.arr[rowx,self.anchor_colx+1:] = df[label].as_matrix()
        return self
                         
    def get_variable_locations_by_row(self, varlist):
        var_to_rows = {}
        column_with_labels = self.arr[:,self.anchor_colx]
        for rowx, label in enumerate(column_with_labels):
            if label in varlist:
                # +1 to rebase from 0  
                var_to_rows[label] = rowx + 1        
        return var_to_rows  
        
    def pop_equations(self):       
        equations = []        
        for label in self.dataset.columns:
            if "=" in label:
                equations.append(label)
                self.dataset = self.dataset.drop(label, 1)
            elif " " in label.strip():
                self.dataset = self.dataset.drop(label, 1)
        return equations       

   
class XlSheet():
    
    def __init__(self, filepath, sheet_n = 1, anchor = 'A1'):
    
        self.input_file_path = filepath
        self.sheet_n = sheet_n

        # warning: will not accept sheet names as strings
        #          may need sheet names to sheet numbers converter
        arr = self.read_sheet_as_array(filepath, sheet_x=sheet_n-1)
        self.image = SheetImage(arr, anchor)
    
    @staticmethod
    def read_sheet_as_array(filename, sheet_x=0):
        """Read sheet_x-th sheet from an Excel file into an numpy's ndarray"""
        contentstring = open(filename, 'rb').read()
        book  = xlrd.open_workbook(file_contents=contentstring)
        sheet = book.sheets()[sheet_x]
        array = np.empty((sheet.nrows,sheet.ncols), dtype=object)
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                value = sheet.cell(row, col).value
                # force type to 'int' where possible
                if isinstance(value, float) and round(value) == value:
                    value = int(value)                
                array[row][col] = value
        return array

    def save(self, filepath=None, sheet=None):

        output_array = self.image.insert_formulas().arr
        
        if not filepath:
            filepath = self.input_file_path
        if not sheet:
            sheet = self.sheet_n
        
        def get_abspath(filepath):      
            folder = os.path.dirname(os.path.abspath(__file__))
            if not os.path.split(filepath)[0]:
                # 'filepath' was file name only  
                return os.path.join(folder, filepath)
            else:
                # 'filepath' was long path
                return filepath
            
        abspath = get_abspath(filepath)        
        wb = Workbook(abspath)
        # later: must check sheet_n exists
        Sheet(sheet).activate()
        Range("A1").value = output_array  
        wb.save()
        return self 


if __name__ == "__main__":
    
    from basefunc import is_equal  

    df1 = XlSheet("xl.xls", sheet_n=1, anchor="A1").image.dataset
    df2 = XlSheet("xl.xls", sheet_n=2, anchor="B3").image.dataset    
    assert is_equal(df2, df1)
    
    xl = XlSheet('xl.xls', 1, "A1")
    arr = xl.image.insert_formulas().arr
    
    XlSheet('xl.xls', 1, "A1").save(sheet=3)
    xl2 = XlSheet('xl.xls', 2, "B3").save(sheet=4)

    def read_range_as_df(filename, sheet_n, anchor):
         return XlSheet(filename, sheet_n, anchor).image.dataset

    df3 = read_range_as_df("xl.xls", sheet_n=3, anchor="A1")
    df4 = read_range_as_df("xl.xls", sheet_n=4, anchor="B3")  
    assert is_equal(df3, df4)