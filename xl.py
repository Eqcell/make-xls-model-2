# -*- coding: utf-8 -*-
"""
Created on Sun May 29 09:12:29 2016

@author: Евгений
"""

import sys
import os

import numpy as np
import pandas as pd

import xlrd
from xlwings import Workbook, Range, Sheet

from basefunc import to_rowcol
from model import MathModel

        
def get_xlrd_sheet(filename, sheet):
   
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
    sheet = get_xlrd_sheet(filename, sheet)       
    array = np.empty((sheet.nrows,sheet.ncols), dtype=object)
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            value = sheet.cell(row, col).value
            # force values type to 'int' where possible
            if isinstance(value, float) and round(value) == value:
                value = int(value)                
            array[row][col] = value
    return array              

def write_array_to_sheet(filepath, sheet, arr):

    def _make_abspath(filepath):      
        folder = os.path.dirname(os.path.abspath(__file__))
        if not os.path.split(filepath)[0]:
            # 'filepath' was file name only  
            return os.path.join(folder, filepath)
        else:
            # 'filepath' was long path
            return filepath
            
    # Workbook(path) seems to fail unless full path is provided
    path = _make_abspath(filepath)
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