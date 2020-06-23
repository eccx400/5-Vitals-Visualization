import json
import xlwt
import glob
from xlwt import Workbook
import xlsxwriter
import matplotlib.pyplot as plt
import win32com.client
from win32com.client import DispatchEx
import pandas as pd
import numpy as np
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from functools import reduce
pd.options.mode.chained_assignment = None  # default='warn'

df = pd.read_csv (r'C:\Users\datiphy\Documents\NEO Excel\46520\subject_16081_ChartEvents.txt', index_col= False, sep= '\t',
                    names= ['ROW_ID', 'SUBJECT_ID', 'HADM_ID', 'ICUSTAY_ID', 'ITEMID', 'CHARTTIME', 'STORETIME', 'CGID', 'VALUE', 'VALUENUM', 'VALUEUOM', 'WARNING', 'ERROR', 'RESULTSTATUS'])

print(df)

out_path = "C:/Users/datiphy/Documents/NEO Excel/Reports/16081CE_Report.xlsx"
writer = pd.ExcelWriter(out_path, engine='xlsxwriter')

df.to_excel( writer, sheet_name='Visualization')

writer.save()