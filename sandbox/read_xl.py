def read_array(filename, sheet):    
    df = read_sheet(filename, sheet, None)    
    return df.fillna("").astype(object).as_matrix().transpose()

def read_sheet(filename, sheet, header):    
    return pd.read_excel(filename, sheetname=sheet, header = header).transpose()