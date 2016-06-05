from basefunc import is_equal
from xl import XlSheet
   
def test_xl_sheet_reading():    
    
    df1 = XlSheet("xl.xls", sheet_n=1, anchor="A1").image.dataset
    df2 = XlSheet("xl.xls", sheet_n=2, anchor="B3").image.dataset    
    assert is_equal(df2, df1)
    
    #xl = XlSheet('xl.xls', 1, "A1")
    #arr = xl.image.insert_formulas().arr
    
def test_xl_sheet_end_to_end():    
    
    XlSheet('xl.xls', 1, "A1").save(sheet=3)
    XlSheet('xl.xls', 2, "B3").save(sheet=4)

    def read_range_as_df(filename, sheet_n, anchor):
         return XlSheet(filename, sheet_n, anchor).image.dataset

    df3 = read_range_as_df("xl.xls", sheet_n=3, anchor="A1")
    df4 = read_range_as_df("xl.xls", sheet_n=4, anchor="B3")  
    assert is_equal(df3, df4)   