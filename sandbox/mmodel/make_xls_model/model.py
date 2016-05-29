"""Make spreadsheet model in Excel file based on historic data, equations, and control parameters.

   Reads inputs from 'data', 'controls', 'equations' and 'names' sheets of <xlfile> and writes 
   resulting spreadsheet to 'model' sheet in <xlfile>. Overwrites 'model' sheet in <xlfile> 
   without warning.  

Usage:
    model.py <xlfile> 
    model.py <xlfile> (--split-dataset | -D)
    model.py <xlfile> (--make-model    | -M) [--slim | -s]
    model.py <xlfile> (--update-model  | -U) [--sheet=<name>]   
"""


# Flags and options:
# --use-dataset or -D  derive 'data', 'controls', 'names' and 'equations' sheets content from 'dataset' sheet
# --slim or -s          produce no extra formatting on 'model' sheet (labels and years only).
# --update-model or -U        update Excel formulas on 'model' sheet or other sheet specified in [--sheet=<name>]


from docopt import docopt
import os
from make_xl_model import make_xl_model, update_xl_model, derive_sheets_from_dataset, ModelCreator, ModelUpdater, DatasetSplitter
from globals import MODEL_SHEET

def get_filepath(arg):
    """Returns absolute path to <xlfile>"""
    return os.path.abspath(arg["<xlfile>"])
    
def get_model_sheet(arg):
    if arg['--sheet'] is not None:
        return arg['--sheet']
    else:
        return MODEL_SHEET


if __name__ == "__main__old":
   
    arg = docopt(__doc__)

    file = get_filepath(arg)
    sheet = get_model_sheet(arg)
    slim = False

    # slim formatting
    if arg["--slim"] or arg["-s"]:
        slim = True

    if arg["-U"] or arg["--update-model"]:
        update_xl_model(file, sheet)
    elif arg["--use-dataset"] or arg["-D"]:
        derive_sheets_from_dataset(file)
    else:
        make_xl_model(file, sheet, slim)


if __name__ == "__main__":

    arg = docopt(__doc__)

    file = get_filepath(arg)
    sheet = get_model_sheet(arg)
    slim = False

    # slim formatting
    if arg["--slim"] or arg["-s"]:
        slim = True

    if arg["-U"] or arg["--update-model"]:
        _ = ModelUpdater(file, sheet)
        _.update_model()
        _.print_model_sheet()
        _.save()
    elif arg["--split-dataset"] or arg["-D"]:
        _ = DatasetSplitter(file)
        _.derive_from_dataset()
        _.save()
    else:
        _ = ModelCreator(file, sheet)
        if slim:
            _.build_slim()
        else:
            _.build_fancy()
        _.print_model_sheet()
        _.save()
