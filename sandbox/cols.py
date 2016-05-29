# -*- coding: utf-8 -*-
"""
Created on Mon May 23 22:32:40 2016

@author: Евгений
"""

def col_to_num(col_str):
    """ Convert base26 column string to number. """
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1

    return col_num
    
assert col_to_num("B") == 2