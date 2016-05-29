# -*- coding: utf-8 -*-
"""
Created on Fri Jul 31 01:37:09 2015

@author: Евгений
"""
from collections import OrderedDict

list_otuples = [("Z", "eqZ"), ("GDP", "eqGDP"), ("ABC", "eqABC"), ("_Var0", 0)]
od = OrderedDict(list_otuples)
for x in od.keys():
    print (x)

no = dict(list_otuples )
for x in no.keys():
    print (x)


z = OrderedDict()
z["GDP"] = "..."
print(z)    