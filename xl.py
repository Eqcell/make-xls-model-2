# -*- coding: utf-8 -*-
"""
Created on Sun May 29 09:12:29 2016

@author: Евгений
"""

import sys

import numpy as np
import pandas as pd

import xlrd
from xlwings import Workbook, Range, Sheet
import os

from basefunc import to_rowcol
from model import MathModel

class SheetImage():
    """
    Numpy array representing cells in Excel sheet, with optional anchor cell like "A1".  
    .insert_formulas() method populates cells in forecast periods with excel formulas.
    """
    
    def __init__(self, arr, anchor):
        """
        Inputs
        ------
        arr : numpy array  
        anchor : string with A1 style reference, defaults to "A1"
        """
    
        self.arr = arr
        self.anchor_rowx, self.anchor_colx = to_rowcol(anchor, base = 0)

        self.dataset = self.extract_dataframe().transpose()
        self.equations = self.pop_equations()
        self.var_to_rows = self.get_variable_locations_by_row(varlist=self.dataset.columns)
        self.model = MathModel(self.dataset, self.equations)\
                     .set_xl_positioning(self.var_to_rows, anchor)        

    def extract_dataframe(self):
        """Return a part of 'self.arr' starting anchor cell as dataframe.""" 
           
        data = self.arr[self.anchor_rowx:,self.anchor_colx:]
        
        return pd.DataFrame(data=data[1:,1:],    # values
                           index=data[1:, 0],    # 1st column as index
                         columns=data[0 ,1:])    # 1st row as the column names

    def insert_formulas(self):
        """Populate formulas on array representing Excel sheet."""        
        df = self.model.get_xl_dataset()
        column_with_labels = self.arr[:,self.anchor_colx]
        for rowx, label in enumerate(column_with_labels):
            if label in df.columns:
                self.arr[rowx,self.anchor_colx+1:] = df[label].as_matrix()
        return self
                         
    def get_variable_locations_by_row(self, varlist):
        """Return a part of 'self.arr' starting anchor cell as dataframe.""" 
        var_to_rows = {}
        column_with_labels = self.arr[:,self.anchor_colx]
        for rowx, label in enumerate(column_with_labels):
            if label in varlist:
                # +1 to rebase from 0  
                var_to_rows[label] = rowx + 1        
        return var_to_rows  
        
    def pop_equations(self):       
        """Return list of strings containing equations. 
           Also cleans self.dataset off junk non-variable columns""" 
        equations = []        
        for label in self.dataset.columns:
            if "=" in label:
                equations.append(label)
                self.dataset = self.dataset.drop(label, 1)
            elif " " in label.strip():
                self.dataset = self.dataset.drop(label, 1)
        return equations       

   
class XlSheet():
    """Access Excel file for reading sheet and saving sheet with formulas.
    
    XlSheet(filename).save() will read first sheet of Excel file and populate it with formulas.
    
    """
    
    def __init__(self, filepath, sheet = 1, anchor = 'A1'):
        """
        Inputs
        ------
        filepath : valid path to Excel file, xls only, xlsx not supported
        sheet: string or integer >=1, representing sheet name or number starting at 1, defaults to first sheet 
        anchor : string with A1 style reference, defaults to "A1"
        """        
    
        self.input_file_path = filepath
        self.input_sheet = sheet

        arr = self.read_sheet_as_array(filepath, sheet)
        self.image = SheetImage(arr, anchor)
    
    @staticmethod
    def get_xlrd_sheet(filename, sheet):
       
       contentstring = open(filename, 'rb').read()
       book = xlrd.open_workbook(file_contents=contentstring)
       
       def to_int(s):
           try:
               return int(s)
           except ValueError:
               return s

       sheet = to_int(sheet)       
       
       if isinstance(sheet, int):
           # if 'sheet' is integer, we assume 'sheet' is based at 1   
           sheet_x = sheet-1
           return book.sheet_by_index(sheet_x)
       elif isinstance(sheet, str) and sheet in book.sheet_names():
           return book.sheet_by_name(sheet)
       else:
           raise Exception("Failed to locate sheet :" + str(sheet))

    @staticmethod
    def read_sheet_as_array(filename, sheet):
        """Read sheet from an Excel file into an numpy's ndarray."""        
        sheet = XlSheet.get_xlrd_sheet(filename, sheet)        
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
            sheet = self.input_sheet
        
        def get_abspath(filepath):      
            folder = os.path.dirname(os.path.abspath(__file__))
            if not os.path.split(filepath)[0]:
                # 'filepath' was file name only  
                return os.path.join(folder, filepath)
            else:
                # 'filepath' was long path
                return filepath
            
        # Workbook(path) seems to fail unless full path is provided
        abspath = get_abspath(filepath)        
        wb = Workbook(abspath)
        # not todo: must check sheet_n exists and create it if not
        Sheet(sheet).activate()
        Range("A1").value = output_array  
        wb.save()
        return self 

def cli():
    """Simple command line interface to XlSheet(filepath, sheet, anchor).save()."""
    if len(sys.argv) == 1:
       raise Exception("Need at least one argument <filename>")  
    elif len(sys.argv) >= 2:
        filename = sys.argv[1]
        xl = XlSheet(filename)
    elif len(sys.argv) == 3:
        filename = sys.argv[1]
        sheet = int(sys.argv[2])
        xl = XlSheet(filename, sheet)
    elif len(sys.argv) == 4:        
        filename = sys.argv[1]
        sheet = int(sys.argv[2])
        anchor = sys.argv[3]
        xl = XlSheet(filename, sheet, anchor)
    xl = xl.save()
    print("Updated formulas in " + filename + ":")
    eqs = ["    " + k + " = " + v  for k, v in xl.image.model.equations.items()]
    for e in eqs:
        print(e)    

if __name__ == "__main__":
    cli()