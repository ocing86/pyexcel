import sys, os, pandas as pd

class ocExcel():
    """Extended class to work with excel file
    """
    #region CLASS FIELDS
    DEFAULT_PATH = os.path.abspath(__file__).replace('ocing86.py','')
    XL_PATH = ''
    FILE_NAME = ''
    DF_DICT = {}
    SH_NAMES = ''
    #endregion

    def readExcel(self):
        """Read Excel File
        Returns:
            [tuple]: [tuple of sheet_names and Excel dictionary {sheet_name : DataFrame}]
        """
        if not os.path.isfile(self.XL_PATH): return FileNotFoundError
        self.SH_NAMES = pd.ExcelFile(self.XL_PATH).sheet_names
        self.DF_DICT = pd.read_excel(self.XL_PATH, sheet_name=self.SH_NAMES)
        return self.SH_NAMES, self.DF_DICT
    
    def readSheet(self, shname:None, hdr:None, idx:None):
        """Read Single Sheet of Excel File
        Returns:
            [DataFrame]
        """
        if not os.path.isfile(self.XL_PATH): return FileNotFoundError
        try:
            if not shname is None and not hdr is None and not idx is None:
                return pd.read_excel(self.XL_PATH, sheet_name=shname, header=hdr, index_col=idx)
        except:
            return Exception(f'Error happened!, *Check if {self.FILE_NAME} exist in directory; * Check if {shname} exist in target workbook')
    
    def getSheet(self, sheet_name):
        try: return self.DF_DICT[sheet_name]
        except: return Exception(f'no such sheet ({sheet_name}) in DataFrame dictionary')

    def getColumn(self, df, column_name): 
        try: return df.loc[:,column_name]
        except: return Exception(f'no such column ({column_name}) in DataFrame')

    def getColumnBySheet(self, sheet_name, column_name): 
        try: return self.getSheet(sheet_name).loc[:,column_name]
        except: return Exception(f'no such column ({column_name}) in DataFrame')

    def getAppendSheet(self):
        dfAppend = pd.DataFrame()
        for df in self.DF_DICT: dfAppend = pd.DataFrame.append(self=dfAppend, other=self.DF_DICT[df])
        return dfAppend

    def __init__(self, fileName='', xlPath=None):
        self.FILE_NAME = fileName
        if xlPath is None:
            self.XL_PATH = self.DEFAULT_PATH + self.FILE_NAME
        else:
            self.XL_PATH = xlPath