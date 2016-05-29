"""Test for make_xls_model"""

import numpy as np
import pandas as pd
import os

import pytest

continuous_years_on_data_and_controls_xl_file = os.path.abspath(os.path.join("examples", "spec.xls"))
assert os.path.exists(continuous_years_on_data_and_controls_xl_file)

all_years_on_controls_xl_file = os.path.abspath(os.path.join("examples", "spec_all_years_on_controls.xls"))
assert os.path.exists(all_years_on_controls_xl_file)

###########################################################################
## Data input
###########################################################################


def test_data_import():    
    from import_specification import get_data_df
    df = get_data_df(continuous_years_on_data_and_controls_xl_file)
    df0 = pd.DataFrame({
            # Note we use float value.0 for GDP, otherwise df.equals(df0) will fail
            "GDP": [66190.0, 71406.0]
          , "GDP_IQ": [101.3407, 100.6404]
          , "GDP_IP": [105.0467, 107.1941]
        },
        index=[2013, 2014]
    )[["GDP", "GDP_IQ", "GDP_IP"]]
    assert df.equals(df0)


def test_controls_import():
    from import_specification import get_controls_df
    df = get_controls_df(continuous_years_on_data_and_controls_xl_file)
    df0 = pd.DataFrame({
            "GDP_IQ": [95.0, 102.5]
          , "GDP_IP": [115.0, 113.0]
        },
        index=[2015, 2016]
    )[["GDP_IQ", "GDP_IP"]]
    assert df.equals(df0)


def test_equation_import():     
    from import_specification import get_equations
    eq_list, eq_dict = get_equations(continuous_years_on_data_and_controls_xl_file)
    assert eq_dict == {'GDP': 'GDP[t-1] * GDP_IQ / 100 * GDP_IP / 100'}

###########################################################################
## Equations tests
########################################################################### 

def test_misspecified_equations():
    pass

def test_eq_with_comment():
    pass

def test_eq_without_eq_signt():
    pass

def test_imported_test():
    from formula_parser import test_make_eq_dict
    test_make_eq_dict()

def test_parse_equation_to_xl_formula():
    from formula_parser import parse_equation_to_xl_formula as pfunc
    
    dict_ = {'credit':10, 'liq_to_credit': 9}    
    assert pfunc('liq_to_credit*credit', dict_, 1) == '=B10*B11'
    
    dict_ = {'GDP': 99}   
    assert pfunc('GDP[t]', dict_, 1) == '=B100'
    assert pfunc('GDP'   , dict_, 1) == '=B100'    
    assert pfunc('GDP[t] * 0.5 + GDP[t-1] * 0.5', dict_, 1) == '=B100*0.5+A100*0.5'
    
    # todo:    
    
    '''
    >>> parse_equation_to_xl_formula('liq[t] + credit[t] * 0.5 + liq_to_credit[t] * 0.5',
    ...                              {'credit': 2, 'liq_to_credit': 3, 'liq': 4}, 1)
    '=B5+B3*0.5+B4*0.5'
    
    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2}, 1)
    '=B2+A3*100'
    
    >>> parse_equation_to_xl_formula('GDP[n] + GDP_IQ[n-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2}, 1)
    '=B2+A3*100'
    
    If some variable is missing from 'variable_dict' raise an exception:
    
    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100', # doctest: +IGNORE_EXCEPTION_DETAIL 
    ...                              {'GDP': 1}, 1)
    Traceback (most recent call last):  
    KeyError: Cannot parse formula, formula contains unknown variable: GDP_IQ
    '''
    
    # If some variable is included in variables_dict but do not appear in formula_string
    # do nothing.
    assert pfunc('GDP[t] + GDP_IQ[t-1] * 100',
           {'GDP': 1, 'GDP_IQ': 2, 'GDP_IP': 3}, 1) == '=B2+A3*100'


###########################################################################
## Misc
###########################################################################
def test_variable_names_validation():
    from import_specification import validate_variable_names
    validate_variable_names([])
    validate_variable_names(['no_dots1', 'no_dots2'])
    with pytest.raises(ValueError) as e:
        validate_variable_names(['no_dots', 'with.dot'])


###########################################################################
## Final result
###########################################################################
    
from make_xl_model import get_resulting_workbook_array_for_make as get_ar      

spec_model_array = np.array([
      ['', '2013', '2014', '2015', '2016']
     ,['GDP', 66190, 71406, '=C2*D3/100*D4/100', '=D2*E3/100*E4/100']
     ,['GDP_IQ', 101.3407,  100.6404, 95.0,  102.5]
     ,['GDP_IP', 105.0467,  107.1941, 115.0, 113.0]
     ,['is_forecast', 0.0, 0.0, 1.0, 1.0]
    ]
    , dtype=object)

def test_resulting_array_spec_xls():
    assert np.array_equal(spec_model_array, get_ar(continuous_years_on_data_and_controls_xl_file, slim=True))
    assert np.array_equal(spec_model_array, get_ar(all_years_on_controls_xl_file, slim=True))