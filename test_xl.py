from basefunc import is_equal
from xl import XlSheet
 
PATH = "test1.xls"
 
def test_xl_sheet_reading():    
    
    df1 = XlSheet(PATH, sheet_n=1, anchor="A1").image.dataset
    df2 = XlSheet(PATH, sheet_n=2, anchor="B3").image.dataset    
    assert is_equal(df2, df1)
    
    #xl = XlSheet('xl.xls', 1, "A1")
    #arr = xl.image.insert_formulas().arr
 
def read_by_name_and_int():
   sh_name = 'input_sheet_v1'
   assert sh_name == XlSheet.get_xlrd_sheet(PATH, 1).name
   assert sh_name == XlSheet.get_xlrd_sheet(PATH, sh_name).name
   
def test_xl_sheet_end_to_end():        
    
    XlSheet(PATH, 1, "A1").save(sheet=3)
    XlSheet(PATH, 2, "B3").save(sheet=4)

    def read_range_as_df(filename, sheet_n, anchor):
         return XlSheet(filename, sheet_n, anchor).image.dataset

    df3 = read_range_as_df(PATH, sheet_n=3, anchor="A1")
    df4 = read_range_as_df(PATH, sheet_n=4, anchor="B3")  
    assert is_equal(df3, df4)   