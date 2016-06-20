import os
from xl import ExcelSheet

   
     
def example_full_path(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'examples', filename)
     
path = example_full_path('bdrn.xls')
xl = ExcelSheet(path, "model", "c1").save()
xl.echo()


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