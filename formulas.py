# -*- coding: utf-8 -*-
"""
Created on Mon May 23 01:22:27 2016

@author: Евгений
"""

from collections import OrderedDict
import xlrd
import re
TIME_INDEX_VARIABLES = ['t']

###########################################################################
## Equations list -> equations dictionary
########################################################################### 


def parse_equation_string(string):
    left_hand_side_expression, formula = string.split('=')
    varname = left_hand_side_expression.replace(" ","").replace("[t]", "")    
    return varname, formula                     

def error_duplicate_equation(key, eq1, eq2):
    raise ValueError("Two equations for the same variable. " + 
                     "\nVariable: " + key +          
                     "\nExisting equation: " + eq1 +
                     "\nAlternative equation: " + eq2)    
                     
def to_equations_dict(equation_strings):
    eq_dict = OrderedDict()
    # disregard comments and strings without '='
    equation_strings = [eq for eq in equation_strings 
                        if "=" in eq and not eq.strip().startswith("#")]
    for eq in equation_strings:
        key, formula = parse_equation_string(eq)
        if key in eq_dict.keys():
            error_duplicate_equation(key, eq_dict[key], formula)                    
        else:
            eq_dict[key] = formula.strip()
    return eq_dict    

#######################################################################
###
#######################################################################    

def parse_equation_to_xl_formula(formula_string, variables_dict, time_period):
    '''Equivalent method of eqcell_core, but with text-based parser

    >>> parse_equation_to_xl_formula('liq_to_credit*credit', {'credit':10, 'liq_to_credit': 9}, 1)
    '=B10*B11'
    
    >>> parse_equation_to_xl_formula('liq_to_credit*credit', {'liq_to_credit': 9, 'credit':10}, 1)
    '=B10*B11'
        
    >>> parse_equation_to_xl_formula('GDP[t]', {'GDP': 99}, 1)
    '=B100'
    
    >>> parse_equation_to_xl_formula('GDP[t] * 0.5 + GDP[t-1] * 0.5',
    ...                              {'GDP': 99}, 1)
    '=B100*0.5+A100*0.5'
    
    >>> parse_equation_to_xl_formula('GDP * 0.5 + GDP[t-1] * 0.5',
    ...                              {'GDP': 99}, 1)
    '=B100*0.5+A100*0.5'

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
    
    If some variable is included in variables_dict but do not appear in formula_string
    do nothing.
    
    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2, 'GDP_IP': 3}, 1)
    '=B2+A3*100'

    '''

    # Strip whitespace
    formula_string = strip_all_whitespace(formula_string)

    # Expands shorthand
    formula_string = expand_shorthand(formula_string, variables_dict.keys())

    # parse and substitute time indices, eg. GDP[t-1] -> GDP[3] if t = 4
    formula_string = substitute_time_indices(formula_string, time_period)
    
    # each setment in var_time_segments  is like 'GDP[0]', 'GDP_IQ[10]', etc
    var_time_segments = re.findall(r'(\w+\[\d+\])', formula_string)

    for segment in var_time_segments:
        formula_string = replace_segment_in_formula(formula_string, segment, variables_dict)

    return '=' + formula_string

def strip_all_whitespace(string):
    return re.sub(r'\s+', '', string)
    
def get_A1_reference(segment, variables_dict):
    var, period = extract_var_time(segment)
    if var in variables_dict.keys():
        cell_row = get_cell_row(var, variables_dict)
        cell_col = period
        return get_excel_ref(cell_row, period)
    else:
        raise KeyError("Cannot parse formula, formula contains unknown variable: " + var)
    
def replace_segment_in_formula(formula_string, segment, variables_dict):
    A1_ref = get_A1_reference(segment, variables_dict)
    # Match beginning of word
    return re.sub(r'\b' + re.escape(segment), A1_ref, formula_string)
    
def get_cell_row(var, variables_dict):
    try:
        return variables_dict[var] # Variable offset in file
    except KeyError:
        raise ValueError('Variable %s is in formula, but not found in variables_dict' % repr(var))

def get_excel_ref(row, col):
    '''
    >>> get_excel_ref(0, 0)
    'A1'
    >>> get_excel_ref(3, 2)
    'C4'
    '''
    return xlrd.colname(col) + str(row)

def substitute_time_indices(formula_string, period):
    '''
    >>> substitute_time_indices('GDP[t]+GDP[t-1]+0.5*GDP_IP[t]', 1)
    'GDP[1]+GDP[0]+0.5*GDP_IP[1]'
    >>> substitute_time_indices('GDP[t]+GDP[n-1]+0.5*GDP_IP[n]', 1)
    'GDP[1]+GDP[0]+0.5*GDP_IP[1]'
    '''
    
    # time index in square brackets [], [<ws><litteral in TIME_INDEX_VARIABLES><ws >+-<ws><integer><ws>]
    TI = ''.join(TIME_INDEX_VARIABLES)
    # note here [] are part of regex notation 
    TI_REGEX = r'[' + TI + r'+\-\d]'
    #TI_REGEX = r'[t+\-\d]'
    
    for time_index in re.findall(r'\[(' + TI_REGEX + '+)\]', formula_string):
        # We transfrom TIME_INDEX_VARIABLES to t for proper evaluation
        period_normalize = re.sub('[' + TI + ']', 't', time_index)
        try:
            t = period
            period_offset = eval(period_normalize)     # evaluate time expression t-1
        except:
            raise ValueError('Time expression %s[%s] invalid' % (var, period))
        
        formula_string = formula_string.replace('[' + time_index + ']', 
                                                      '[' + str(period_offset) + ']')
    return formula_string

def expand_shorthand(formula_string, variables):
    """
    >>> expand_shorthand('GDP_IQ+GDP_IP+GDP_IQ[t-1]', ['GDP', 'GDP_IP', 'GDP_IQ'])
    'GDP_IQ[t]+GDP_IP[t]+GDP_IQ[t-1]'
    
    >>> expand_shorthand('GDP * 0 + GDP [t-1] * GDP_IQ / 100 * GDP_IP[t] / 100', 
    ...                        ['GDP', 'GDP_IP', 'GDP_IQ'])
    'GDP[t] * 0 + GDP [t-1] * GDP_IQ[t] / 100 * GDP_IP[t] / 100'
    
    >>> expand_shorthand('GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100',
    ...                  ['GDP', 'GDP_IP', 'GDP_IQ'])
    'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100'
    
    >>> expand_shorthand('liq_to_credit*credit', ['liq_to_credit', 'credit'])
    'liq_to_credit[t]*credit[t]'
    
    >>> expand_shorthand('liq_to_credit*credit', ['credit', 'liq_to_credit'])
    'liq_to_credit[t]*credit[t]'
    
    """
    for var in variables:
        if var != '':
            formula_string = re.sub(var + r'(?!\s*[\dA-Za-z_^\[])',
                                       var + '[t]', formula_string) 
    return formula_string


def extract_var_time(formula_string):
    '''Extract variable and time period from formula segment
    
    >>> extract_var_time('GDP[1]')
    ('GDP', 1)
    '''
    # Extract group (GDP, 0)
    a, b = re.search(r'(\w+)\[(\d+)\]', formula_string).groups()
    return a, int(b)