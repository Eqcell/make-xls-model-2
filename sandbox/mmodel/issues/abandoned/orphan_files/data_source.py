# coding: utf-8
"""
Data sources for the model and output.
Current entry point: 
    model_spec, view_spec = get_specification()
    model_spec, view_spec = get_mock_specification(user_input)
    
"""
from pprint import pprint
import pandas as pd
import numpy as np


###########################################################################
## Sample (mock) proxies as func and constants - to use in this file
###########################################################################

    

# label, year, value
DATA_PROXY = [ ("GDP", 2013, 66190.11992)
        , ("GDP",    2014, 71406.3992)
        , ("GDP_IQ", 2013, 101.3407976)
        , ("GDP_IQ", 2014, 100.6404858)
        , ("GDP_IP", 2013, 105.0467483)
        , ("GDP_IP", 2014, 107.1941886) ] 
        
DATA_PROXY_AS_DF = pd.DataFrame(
             { "GDP": [66190.11992, 71406.3992 ]
          , "GDP_IQ": [101.3407976, 100.6404858]       
          , "GDP_IP": [105.0467483, 107.1941886]}
          , index = [2013, 2014])
          #[["GDP", "GDP_IQ", "GDP_IP"]]
    
# label, year, value
CONTROLS_PROXY = [("GDP_IQ", 2015, 95.0)
        , ("GDP_IP", 2015, 115.0)
        , ("GDP_IQ", 2016, 102.5)
        , ("GDP_IP", 2016, 113.0)
        , ("is_forecast", 2015, 1)
        , ("is_forecast", 2016, 1)
        ]        
        
# title, label, group, level, precision
# ERROR: wont print cyrillic charactes, only whitespace.
NAMES_CSV_PROXY = [("ВВП",                      "GDP",    "Нацсчета", 1, 0),
                   ("Индекс физ.объема ВВП",    "GDP_IQ", "Нацсчета", 2, 1),
                   ("Дефлятор ВВП",	            "GDP_IP", "Нацсчета", 2, 1)]
 
EQ_SAMPLE = ["GDP(t) = GDP(t-1) * GDP_IP(t) / 100 * GDP_IQ(t) / 100"]

# change in test setting: one variable not in output 
ROW_LABELS_IN_OUTPUT = ["GDP", "GDP_IQ", "GDP_IP"] # , "is_forecast"]

##########################################################################
## Sample (mock) proxies as func - to import outside this file 
###########################################################################

def _sample_for_xfill_dataframe_before_equations():    
    z = pd.DataFrame(
          {   "GDP" : [66190.11992, 71406.3992, None, None]
          , "GDP_IQ": [101.3407976, 100.6404858, 95.0, 102.5]       
          , "GDP_IP": [105.0467483, 107.1941886, 115.0, 113.0]          
          , "is_forecast": [None, None, 1, 1]} 
          ,   index = [2013, 2014, 2015, 2016]          
          )
          
    # Test setting: dataframe before equations has less columns than union of controls and data
    return z[ROW_LABELS_IN_OUTPUT]

def _sample_for_xfill_array_after_equations():
    return np.array(   
    [['', '2013', '2014', '2015', '2016']
    ,['GDP', 66190.11992, 71406.3992, '=C2*D3*D4/10000', '=D2*E3*E4/10000']
    ,['GDP_IQ', 101.3407976, 100.6404858, 95.0, 102.5]
    ,['GDP_IP', 105.0467483, 107.1941886, 115.0, 113.0]
    #,['is_forecast', "", "", 1, 1]
    ]
    , dtype=object)    
    # WARNING: actual intention was '=C2*D3/100*D4/100', '=C2*D3/100*D4/100'
    
###########################################################################
## Entry points
###########################################################################

def get_proxy_specification_dict():
    return       {'data': convert_tuple_to_df(DATA_PROXY),     
              'controls': convert_tuple_to_df(CONTROLS_PROXY),
             'equations': EQ_SAMPLE,
                'format': ROW_LABELS_IN_OUTPUT }

# WARNING: to de dereciated in favor of  get_proxy_specification_dict()
def get_mock_specification():
    model_spec = [
    ("Historic data as df",       convert_tuple_to_df(DATA_PROXY)      ),
    ("Names as dict",             {x[1]:x[0] for x in NAMES_CSV_PROXY} ),
    ("Equations as list",         EQ_SAMPLE                            ),
    ("Control parameters as df",  convert_tuple_to_df(CONTROLS_PROXY)  )] 
    
    # LATER: requires workaround
    view_spec = [
    ['Excel filename' ,    'model.xls'],
    ['Sheet name' ,        'model'],
    ['List of variables',  ROW_LABELS_IN_OUTPUT] 
    ]
    
    return model_spec, view_spec

     
def print_specification(specification):             
   for spec in specification:
       print("\n------ {}:".format(spec[0]))
       pprint(spec[1])

###########################################################################
## General handling
###########################################################################
        
def convert_tuple_to_df(tuple_):
    """Returns a dataframe with years in rows and variables in columns. 
       *lt* is a list of tuples like *data_proxy* and *controls_proxy*"""  
    
    # Read dataframe
    df = pd.DataFrame(tuple_, columns=['prop', 'time', 'val'])
    # Pivot by time
    return df.pivot(index='time', columns='prop', values='val')
        
###########################################################################
## Historic data 
###########################################################################
        
def check_get_historic_data_as_dataframe():
    """
    >>> check_get_historic_data_as_dataframe()
    True
    """
    df1 = convert_tuple_to_df(DATA_PROXY)
    df2 = DATA_PROXY_AS_DF  
    return df1.equals(df2)

if __name__ == "__main__":
    import doctest
    doctest.testmod()
    
    # m, v = get_mock_specification()
    # print_specification(m)                 
    # print_specification(v)                 
