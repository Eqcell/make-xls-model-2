# Problem description

In financial analysis and economic forecasting there is a common type of 'spreadsheet models' in Excel which include the following:
- there is some observed historic data for time series (e.g. balance sheet items); 
- forecast is made by assigning future values to some control parameters (growth rates, elasticities, ratios, etc);
- there are that equations link control parameters to the rest of the variables. 

Large Excel files of this kind often become a mess: 
- the whole picture of equations cannot be seen easily
- cannot guarantee it is the same equation across all cells in row/column 
- cannot replicate or amend many formulas in Excel file fast
- control parameters may be hidden somewhere and it is unclear what really governs your forecast.

This problem grows bigger with your file size, model complexity and number of people working on it. However, we still use Excel for this because it has a great user interface, people can experiment with their own changes quickly, can share a model as one file with no extra dependencies.  

#More comments

The resulting file of ```make-xls-model``` should look the same as if I worked in Excel only - no extra dependecies or VBA code, just a regular stand-alone Excel file with proper formulas in cells.

With spreadsheet models of about 20-50 or more equations I assume there should be a big productivity gain, espacially if model structure is sometimes reviewed. 
