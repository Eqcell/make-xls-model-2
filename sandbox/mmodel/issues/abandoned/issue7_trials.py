import pandas as pd

def _internal_get_dataframe_before_equations():    
    """
       This is a dataframe obtained by merging historic data and future values of control variables.
       WARNING: currently returns a stub. 
    """
    return pd.DataFrame(
          {   "GDP" : [66190.11992, 71406.3992, None, None]
          , "GDP_IQ": [101.3407976, 100.6404858, 95.0, 102.5]       
          , "GDP_IP": [105.0467483, 107.1941886, 115.0, 113.0]
          , "is_forecast": [None, None, 1, 1] } 
          ,   index = [2013, 2014, 2015, 2016]
          )
          
z = _internal_get_dataframe_before_equations()

#this is a permutation of existing columns, it will work
FULL_COL_LIST_REARRANGED = ["GDP", "is_forecast", "GDP_IQ", "GDP_IP"]
SHORTER_COL_LIST = ["GDP", "GDP_IQ", "GDP_IP"]

z.columns = ["GDP", "is_forecast", "GDP_IQ", "GDP_IP"]
print(z)        

#this is a selection of columns, it will not work if there are fewer columns than in original df
try:
   z.columns = ["GDP", "GDP_IQ", "GDP_IP"]
except:
   print ("Error handling dataframe")
   
# this is a selection of columns from 'united' dataframe
print(z[SHORTER_COL_LIST ])
          
