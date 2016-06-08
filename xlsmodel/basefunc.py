import re

import xlrd


def to_xl_ref(row, col, base=1):
    if base == 1:
        return xlrd.colname(col - 1) + str(row)
    elif base == 0:
        return xlrd.colname(col) + str(row + 1)


def col_to_num(col_str):
    """ Convert base26 column string to number. """
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num


def to_rowcol(xl_ref, base=1):
    xl_ref = xl_ref.upper() 
    letters, b = re.search(r'(\D+)(\d+)', xl_ref).groups()
    return int(b) + (base - 1), col_to_num(letters) + (base - 1)


def is_equal(df1, df2):
    # in numpy/pandas nan == nan is False, must substitute nans to compare frames
    # also 1 == 1.0 is false
    # below will only compare identically-labeled DataFrame objects, exceptions if different
    # rows of columns
    flag = df1.fillna("") == df2.fillna("")
    return flag.all().all()
