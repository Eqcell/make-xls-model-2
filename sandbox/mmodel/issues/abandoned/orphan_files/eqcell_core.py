from sympy import var
from config import TIME_INDEX_VARIABLES
import xlrd
import re


############## same code as in parser_function.py ##################

def expand_shorthand(formula_string, variables):
    """
    >>> expand_shorthand('GDP_IQ+GDP_IP+GDP_IQ[t-1]', {'GDP_IP': 1, 'GDP_IQ': 2})
    'GDP_IQ[t]+GDP_IP[t]+GDP_IQ[t-1]'

    >>> expand_shorthand('GDP * 0 + GDP [t-1] * GDP_IQ / 100 * GDP_IP[t] / 100', {'GDP_IP': 1, 'GDP_IQ': 2, 'GDP':3})
    'GDP[t] * 0 + GDP [t-1] * GDP_IQ[t] / 100 * GDP_IP[t] / 100'

    >>> expand_shorthand('GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100', {'': 0, 'GDP_IQ': 2, 'GDP': 1, 'GDP_IP': 3})
    'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100'
    """
    for var in variables:
        if var != '':
            formula_string = re.sub(var + r'(?!\s*[\dA-Za-z_^\[])',
                                       var + '[t]', formula_string)
    return formula_string

def get_excel_ref(cell):
    """
    TODO: test below fails with strange message, need to fix

    >>> get_excel_ref((0,0))
    'A1'

    >>> get_excel_ref((1,3))
    'D2'
    """
    row = cell[0]
    col = cell[1]
    return xlrd.colname(col) + str(row + 1)

def strip_all_whitespace(string):
    return re.sub(r'\s+', '', string)

############## end of imported code ##################

def get_cell_row(variables_dict, var):
    var = str(var)
    try:
        return variables_dict[var] # Variable offset in file
    except KeyError:
        raise KeyError('Variable %s is in formula, but not found in variables_dict' % repr(var))

def convert_brackets(string):
    """
    >>> convert_brackets("GDP[t]")
    'GDP(t)'
    """
    string = string.replace("[", "(")
    string = string.replace("]", ")")
    return string

def check_parse_equation_to_xl_formula():
    """
    >>> check_parse_equation_to_xl_formula()
    =D2*E3*E4/10000
    """
    string_formula = 'GDP[t-1] * GDP_IP[t] / 100 * GDP_IQ[t] / 100'

    # WARNING = actual dict_variables contains {'': 0,}
    dict_variables = {'GDP': 1, 'GDP_IP': 2, 'GDP_IQ': 3}

    print (parse_equation_to_xl_formula(string_formula, dict_variables, 4))

def parse_equation_to_xl_formula(formula_string, variables_dict, time_period):
    '''
    Tests (as in formula_parser.py):

    >>> parse_equation_to_xl_formula('GDP[t]', {'GDP': 99}, 1)
    '=B100'

    >>> parse_equation_to_xl_formula('GDP[t] * 0.5 + GDP[t-1] * 0.5',
    ...                              {'GDP': 99}, 1)
    '=0.5*A100 + 0.5*B100'

    >>> parse_equation_to_xl_formula('GDP * 0.5 + GDP[t-1] * 0.5',
    ...                              {'GDP': 99}, 1)
    '=0.5*A100 + 0.5*B100'

    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2}, 1)
    '=100*A3 + B2'

    >>> parse_equation_to_xl_formula('GDP[n] + GDP_IQ[n-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2}, 1)
    '=100*A3 + B2'

    If some variable is missing from 'variable_dict' raise an exception:

    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100', # doctest: +IGNORE_EXCEPTION_DETAIL
    ...                              {'GDP': 1}, 1)
    Traceback (most recent call last):
    KeyError: Cannot parse formula, formula contains unknown variable: GDP_IQ

    If some variable is included in variables_dict but do not appear in formula_string
    do nothing.

    >>> parse_equation_to_xl_formula('GDP[t] + GDP_IQ[t-1] * 100',
    ...                              {'GDP': 1, 'GDP_IQ': 2, 'GDP_IP': 3}, 1)
    '=100*A3 + B2'

    '''
    # TODO: from this line below
    #       - see if docstrings make sense, correct if necessary
    #       - suggest simplfications, if any

    formula_string = strip_all_whitespace(formula_string)
    formula_string = expand_shorthand(formula_string, variables_dict.keys())
    formula_string = convert_brackets(formula_string)

    varirable_list = [x for x in variables_dict.keys()] + TIME_INDEX_VARIABLES

    # declares sympy variables
    var(' '.join(varirable_list))

    right_side_expression = sympyfy_formula(formula_string)
    return get_excel_formula_string(right_side_expression, time_period, variables_dict)


def sympyfy_formula(string):
    '''
    Convert the formula as a string to the equivalent sympy expression.
    '''
    try:
        return eval(string)     # converting the formula into sympy expressions
    except NameError:
        raise NameError('Undefined variables in formulas')


def get_excel_formula_string(right_side_expression, time_period, variables):
    """
    Using the right-hand side of a math expression (e.g. a(t)=a(t-1)*a_rate(t)),
    converted to SymPy expression, substitute the time index variable (t) in it
    using time_period. The function finds and return the Excel formula corresponding to
    the right-hand side expression as a string.

    input
    -----
    right_side_expression:  SymPy expression, e.g. a(t-1)*a_rate(t)
    time_period:            A value to be substitute for the time index, t.
    variables:              A dictionary that maps variable names to excel row numbers.

    output
    ------
    formula_string:         a string of excel formula, e.g. '=A20*B21'
    """
    right_dict = simplify_expression(right_side_expression, time_period, variables)
    for right_key, right_coords in right_dict.items():
        #excel_index = str(Range(get_sheet(), tuple(right_coords)).get_address(False, False))
        excel_index = get_excel_ref(tuple(right_coords))
        right_side_expression = right_side_expression.subs(right_key, excel_index)
    formula_str = '=' + str(right_side_expression)
    return formula_str


def simplify_expression(expression, time_period, variables, depth=0):
    """
    A recursive function that breaks a SymPy expression into segments,
    where each segment points to one cell on the excel sheet upon substitution
    of time index variable (t). Returns a dictionary of such segments and the computed
    cells.

    input
    -----
    expression:       Sympy expression, e.g: a(t - 1)*a_rate(t)
    time_period:      A value to be substitute for the time index, t.
    variables:        A dictionary that maps variable names to excel row numbers.
    depth:            Depth of recursion, used internally

    output
    ------
    result:           A dict with a segment as key and computed excel cell index as value,
                      e.g: {a(t - 1): (5, 4), a_rate(t): (4, 5)}
    """
    result = {}

    # get the function from sympy expression, e.g for expression = f(t), `f` is the function
    variable = expression.func

    if variable.is_Function:
        # for simple expressions like f(t), variable=f and variable.is_Function = True,
        # for more complex expressions, variable would be another expression, hence would have to be broken down recursively.
        # get the row index from variable name
        # cell_row = variables[str(variable)]
        cell_row = get_cell_row(variables, variable)
        # get the independent var, mostly `t` from the argument in expression
        x = list(expression.args[0].free_symbols)[0]
        cell_col = int(expression.args[0].subs(x, time_period))
        result[expression] = (cell_row, cell_col)
    else:
        if depth > 5:
            raise ValueError("Expression is too complicated: " + expression)

        depth += 1
        for segment in expression.args:
            result.update(simplify_expression(segment, time_period, variables, depth))

    return result

if __name__ == "__main__":
    import doctest
    doctest.testmod()
    pass
