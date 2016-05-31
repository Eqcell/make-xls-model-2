Requirements
------------
 - Windows machine with Microsoft Excel
 - [Anaconda](https://www.continuum.io/downloads#_windows) package suggested for libraries
 - Python 3.5 

User story
----------
  - the user wants to automate filling formulas on Excel sheet
  - Excel sheet has simple 'roll forward' forecast - a spreadsheet model with some 
    historic values for variables and some control parameters for forecast values (e.g. rates of growth) 
  - equations link control parameters and variables previous' period historic values to forecast values
  - equations are written down in excel sheet as text strings like ```y = y[t-1] * rog```

The benefits
------------
  - all formulas for spreadsheet model are written down explicitly as visible text and not just hidden in cells
  - formulas in cells are filled in the same way the user could have done it, resulting file has no extra dependencies
  - one can easily see all control variables that govern the forecast

Terms used
----------
- spreadsheet model
- 'roll-forward' forecast
- equation - formulas like ```y = y[t-1] * rog```
- control variables, controls - variables on right-hand side of equations, which do no appear on left side (e.g ```rog```)
- dependent variables, dependents - variables on the left-hand side of equations (e.g ```rog```)


Rules/requirements
------------------
 - time series in rows only, dataset has horizontal orientation 
 - data range starts next to variable labels and time labels
 - all control variables must be supplied on sheet
 - 'is_forecast' variable required in dataset, it is 0 for historic periods and 1 for forecast periods
 - '[t]' is reserved for indeces
 -  time index for left hand-side variable is always [t] (not [t+1]) 
  
Limitations
-----------
- one sheet only, no multi-sheet models supported
- variable appears on sheet only once
- no equations for historic variables

 
What the script does 
--------------------
- fill cells in Excel sheet with formulas (e.g. '=C3*D4') based on list of variable names and equations
- formulas go only to forecast periods columns (where is_forecast == 1) 

Example
-------

Excel file ```xl.xls``` has following contents:

```
Input Excel sheet:
-------------------------------
            A     B     C     D
1        year  2014  2015  2016
2 is_forecast     0     0     1
3           y    85   100   
4         rog              1.05
6 y = y[t-1] * rog
--------------------------------

Output Excel sheet:
-------------------------------
            A     B     C     D
1        year  2014  2015  2016
2 is_forecast     0     0     1
3           y    85   100  =C3*D4 
4         rog              1.05
6 y = y[t-1] * rog
--------------------------------
```

Comments:
- 'year' is time label, it is not used in calculations 
- 'is_forecast' denotes forecast time periods, it is 0 for historic periods, 1 for forecasted
- 'y' is data variable
- 'rog' (rate of growth) is control parameter
- 'y = y[t-1] * rog' is formula (equation)
