# -*- coding: utf-8 -*-
"""
Created on Sun May 22 14:52:44 2016

@author: Евгений
"""

import pandas as pd

# requirement: years in first row
# region starts at A1

class InputSheet():
    
    def __init__(self, filepath, sheet = 0): 
        
        self.image_df = pd.read_excel(filepath, sheet) 
        
        self.dataset_df = self.image_df.transpose()
        assert 'is_forecast' in self.dataset_df.columns   
        self.eqs = []
        
        for label in self.dataset_df.columns:
            if "=" in label:
                self.eqs.append(label)
                self.dataset_df = self.dataset_df.drop(label, 1)
            elif " " in label.strip():
                self.dataset_df = self.dataset_df.drop(label, 1)
        
        # +1 for rebasing form 0 and +1 for            
        self.var_to_rows = {l:i+1+1 for i, l in enumerate(self.image_df.index)}
        varnames = self.dataset_df.columns
        self.var_to_rows = {k:self.var_to_rows[k] for k in varnames}
        del self.var_to_rows['is_forecast']
        
    def get_dataset(self):
        return self.dataset_df
        
    def get_equations(self):
        return self.eqs    
        
    def get_var_rows(self):
        return self.var_to_rows

if __name__ == "__main__":
    
    ms = InputSheet("xl.xls")
    assert ms.get_dataset().to_csv() == ',is_forecast,y,rog\n2014,0.0,85.0,\n2015,0.0,100.0,\n2016,1.0,,1.05\n'
    assert ms.get_equations() == ['y = y[t-1] * rog']
    assert ms.get_var_rows() == {'y': 3, 'rog': 4}

from xlwings import Workbook, Range, Sheet
import numpy as np

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

def save_df(df, file, sheet):

   wb = Workbook(file)
   Sheet(sheet).activate()
   pm = to_matrix(df)
   Range(sheet, 'A1').value = pm
   wb.save()    

df = ms.get_dataset()
save_df(df.transpose(), file = "D:/git/eqcell2/xl_out.xls", sheet = 1)

def write_array_to_xl_using_xlwings(ar, file, sheet):
    # Note: if file is opened In Excel, it must be first saved before writing 
    #       new output to it, but it may be left open in Excel application. 
    wb = Workbook(file)
    Sheet(sheet).activate()

    def nan_to_empty_str(x):
        return '' if type(x) == float and np.isnan(x) else x

    Range(sheet, 'A1').value = [[nan_to_empty_str(x) for x in row] for row in ar]
    wb.save()
    
    