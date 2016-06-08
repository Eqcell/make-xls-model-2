import os
from xl import XlSheet
     
from model import Formula


# PROBLEM 1
assert 'FondOT[t]+FondOther[t]' == Formula.expand_shorthand("FondOT+FondOther", {"FondOT":0,"FondOther":1})
     
     
# PROBLEM 2     
     
d = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(d, 'examples', 'bdrn.xls')
#xl = XlSheet(path, "model", "c1").save()
arr = XlSheet.read_sheet_as_array(path, "model")
xl = XlSheet(path, "model", "c1")
xl.save()

# PROBLEM 3
# code above passes until .save, however, command line fails differently!!!

#D:\git\make-xls-model-2>python xl.py examples\bdrn.xls model c1
#Traceback (most recent call last):
#  File "xl.py", line 193, in <module>
#    cli()
#  File "xl.py", line 176, in cli
#    xl = XlSheet(filename)
#  File "xl.py", line 104, in __init__
#    self.image = SheetImage(arr, anchor)
#  File "xl.py", line 40, in __init__
#    self.model = MathModel(self.dataset, self.equations)\
#  File "D:\git\make-xls-model-2\model.py", line 213, in __init__
#    self._validate_math_model()
#  File "D:\git\make-xls-model-2\model.py", line 217, in _validate_math_model
#    assert 'is_forecast' in self.dataset.columns
#AssertionError
#"""