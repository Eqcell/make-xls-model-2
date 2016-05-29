Requirements
------------
 - Windows machine with Microsoft Excel
 - Anaconda package suggested for libraries
 - Python 3.5 

Description
-----------

User story: 
  - the user wants to automate filling formulas on Excel sheet 
  - Excel sheet has simple 'roll forward' forecast - a spreadsheet model with some 
    historic variables and some control parameters for forecast (e.g. rates of growth) 
  - equations link control variables and previous period historic values to forecast values
  - equations are written down in excel sheet as text strings like 'y = y[t-1] * rog'
  - by running the python script the user has formulas filled in the Excel where necessary
  - the benefit is to have all model's formulas written down explicitly and not hidden in cells
  - currently we read input Excel sheet and write output sheet to different file or different sheet,
    but mya also write to same sheet to fill formulas
  
 
Some rules: 
  - from equations we know which variables are 'depenendent'('left-hand side')
  - control parameters are right-hand side variables, which do no appear on left side
  - all control variables must be supplied on sheet in dataset
  - we need explicit specification of year when the forecast starts -  by 'is_forecast' vector 
      
Simplifications/requirements:
  - critical, but not checked: 
     - time series in rows only, horizontal orientation 
     - dataset starts at A1 cell
  - checked:
     - must have 'is_forecast' vector in dataset
  - not critical:
     - datablock is next to variable labels
     - time labels are years, not checked for continuity

Main functionality: 
- fill cells in Excel sheet with formulas (e.g. '=C3*D4') based on 
                    list of variable names and equations.
- formulas go only to forecast periods columns (where is_forecast == 1) 

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

Comment:
- 'year' is not used in calculations 
- 'is_forecast' denotes forecast time periods, it is 0 for historic periods, 1 for forecasted
- 'y' is data variable
- 'rog' is control parameter
- 'y = y[t-1] * rog' is formula (equation)
