'''from openpyxl import Workbook

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "hello"
sheet["B1"] = "world!"

workbook.save(filename="hello_world.xlsx")
'''

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

df = pd.read_csv(r'C:\Users\datiphy\Documents\NEO Excel\46520_P\subject_719_prescriptions.csv', sep= '\t',
                    names= ['Prescription_ID', 'ROW_ID', 'SUBJECT_ID', 'HADM_ID', 'ICUSTAY_ID', 'STARTDATE', 'ENDDATE', 'DRUG_TYPE', 'DRUG', 'DRUG_NAME_POE', 'DRUG_NAME_GENERIC', 'FORMULARY_DRUG_CD', 'GSN', 'NDC', 'PROD_STRENGTH', 'DOSE_VAL_RX', 'DOSE_UNIT_RX', 'FORM_VAL_DISP', 'FORM_UNIT_DISP', 'ROUTE'])

print(df)

out_path = "C:/Users/datiphy/Documents/NEO Excel/Reports/subject_719_prescriptions_Report.xlsx"
writer = pd.ExcelWriter(out_path, engine='xlsxwriter')

df.to_excel( writer, sheet_name='Visualization')

writer.save()