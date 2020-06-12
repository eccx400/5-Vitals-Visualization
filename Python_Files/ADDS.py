import xlwt
from xlwt import Workbook
import xlsxwriter
import pandas as pd
import os
import glob
from pandas import ExcelWriter
from pandas import ExcelFile
import xlwings as xw
from win32com.client import Dispatch
from functools import reduce
from StyleFrame import StyleFrame, utils
pd.options.mode.chained_assignment = None  # default='warn'

path1 = 'C:\\Users\\datiphy\\Documents\\NEO Excel\\Charts\\ADDSv3.xlsm'

for filename in glob.glob('C:\\Users\\datiphy\\Documents\\NEO Excel\\TestReport\\*.xlsx'):
    print(filename)
    xl = Dispatch("Excel.Application")

    wb1 = xl.Workbooks.Open(path1)
    wb2 = xl.Workbooks.Open(filename)

    ws1 = wb1.Worksheets(1)
    ws1.Copy(Before=wb2.Worksheets(1))

    wb2.Close(SaveChanges=True)
xl.Quit()