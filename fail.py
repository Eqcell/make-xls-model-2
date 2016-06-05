import os
from xl import XlSheet
     
path = os.path.join('examples', 'bdrn.xls')
XlSheet(path, "model", "c1").save()

# this code above fails as below
"""
D:\git\make-xls-model-2>python fail.py
Traceback (most recent call last):
  File "fail.py", line 5, in <module>
    XlSheet(path, "model", "c1").save()
  File "D:\git\make-xls-model-2\xl.py", line 104, in __init__
    self.image = SheetImage(arr, anchor)
  File "D:\git\make-xls-model-2\xl.py", line 37, in __init__
    self.dataset = self.extract_dataframe().transpose()
  File "D:\git\make-xls-model-2\xl.py", line 49, in extract_dataframe
    index=data[1:, 0],    # 1st column as index
IndexError: index 0 is out of bounds for axis 1 with size 0
"""

# however, command line fails differently!!!
# need to fix both + introduce more py.tests that highlight solutions to occurred errors  
"""
D:\git\make-xls-model-2>python xl.py examples\bdrn.xls model c1
Traceback (most recent call last):
  File "xl.py", line 193, in <module>
    cli()
  File "xl.py", line 176, in cli
    xl = XlSheet(filename)
  File "xl.py", line 104, in __init__
    self.image = SheetImage(arr, anchor)
  File "xl.py", line 40, in __init__
    self.model = MathModel(self.dataset, self.equations)\
  File "D:\git\make-xls-model-2\model.py", line 213, in __init__
    self._validate_math_model()
  File "D:\git\make-xls-model-2\model.py", line 217, in _validate_math_model
    assert 'is_forecast' in self.dataset.columns
AssertionError
"""