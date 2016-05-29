import yaml as ya
from pprint import pprint

# Parse YAML configuration file in Python
# As described in yaml.py
# Max. USD 30 (coding + debugging), 1 day

"""
Usage:
   python yaml_parser.py <YAML_FILENAME>
   python yaml_parser.py <SPECIFICATION_XLS_FILE> <OUTPUT_XLS_FILE>
"""


"""
Rules of YAML parsing for make-xls-model:
-----------------------------------------
A document contains two parts:
    Model specification:
    Output specification:

From Model specification we obtain information about four Excel worksheets,
which can be in one or different Excel files. Sections:
	names of variables:
	historic data:
	control variables:
	formulas:
High level section with 'filename', 'directory' and 'sheet name' is taken as default
for worksheets above.
	filename   : spec.xls
	directory  : D:/data/ # may be omitted in any group
	sheet name : full_spec
	
If no directory is supplied, use current project directory. 
If directory supplied in higher level group - make it default for all subgroups unless overridden.
Values in sample below are defaults, except for mandatory fields and directories.

Minimum valid cofig file:
-------
Model specification:
	filename   : spec.xls                    # MANDATORY, unless speifcied in all 4 subsections below
	
Output specification:
    output file : model.xls                  # MANDATORY 
...

Additional requirements:
- accept backslash in dirnames "D:\path\to" (if possible).	
- check if files exist in 'Model specification:'
- do not check if worksheets exist
- not todo: lists for data sources or lists of sheets of output file/sample file lists
- convert tabs to spaces before parsing to allow tabs in input file 
- raise error if non-latin symbol is in col names ('variable names column', 'formulas column', 'var descriptions column')
- script to be called as python 'yaml_parser.py YAML_FILENAME' or 'yaml_parser.py SPECIFICATION_XLS_FILE OUTPUT_XLS_FILE'
- print message wherever appling defaults
- cmment out parts where any risks may occur. 

Final output:	
- must produce two dictinaries 'model_user_param_dict' and 'output_user_param_dict' as below.
"""     

example_dict = ya.load("""
Model specification:
	filename   : spec.xls                    # MANDATORY, unless speifcied in all 4 subsections below
	directory  : D:/data/
	sheet name : full_spec
	names of variables:
		filename   : spec.xls
		directory  : D:/data/
		sheet name : full_spec
		var descriptions column : A 
		variable names column   : B
	historic data:
		filename   : spec.xls
		directory  : D:/data/  
		sheet name : full_spec
		variable names column : B
		year row              : 2
	control variables:
		filename   : spec.xls
		directory  : D:/data/
		sheet name : full_spec
		variable names column : B
		year row              : 2
	formulas:
		filename   : spec.xls
		directory  : D:/data/
		sheet name : full_spec
		formulas column: B

Output specification:
    output file : model.xls               # MANDATORY 
	directory   : D:/results/             
    sheet:      : model                   # default: same as 'output file' basename
	sample file :                         # optional group, raise warning error if both 'sample file' and 'row labels' provided
		filename   : model_sample.xls
		directory  : D:/results/
		sheet name : model
		variable names column : B
		year row              : 2
    start year : 1995                     
	row labels :                          
	- GDP
	- GDP_IQ     	
	- GDP_IP	
	""".replace("\t", "    "))


model_user_param_dict =  {'control variables': {'filepath': 'D:/data/spec.xls',                                               
                                        'sheet name': 'full_spec',
                             'variable names column': 'B',
                                          'year row': 2},
                  
                         'formulas': {'filepath': 'D:/data/spec.xls'
                                      'formulas column': 'B',
                                      'sheet name': 'full_spec'},
                         'historic data': {'filepath': 'D:/data/spec.xls',
                                           'sheet name': 'full_spec',
                                           'variable names column': 'B',
                                           'year row': 2},
                         'names of variables': {'filepath': 'D:/data/spec.xls',
                                                'sheet name': 'full_spec',
                                                'var descriptions column': 'A',
                                                'variable names column': 'B'}}

output_user_param_dict = {'filepath'   : 'D:/results/model_sample.xls',
                          'row labels' : ['GDP', 'GDP_IQ', 'GDP_IP'],
                          'start year' : 1995}


pprint(example_dict)
pprint(model_user_param_dict) 
pprint(output_user_param_dict)

def get_user_param(yaml_filename):
    """Interface entry point. Returns two dictionaries: model parameters and output parameters
	"""
    output_user_param_dict = {}
	model_user_param_dict = {}
	return  model_user_param_dict, output_user_param_dict