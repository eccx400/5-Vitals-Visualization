import json
import xlwt
import glob
from xlwt import Workbook
import xlsxwriter
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from functools import reduce
pd.options.mode.chained_assignment = None  # default='warn'

__author__ = "Eric Cheng"

files = glob.glob(r'C:\Users\datiphy\Documents\NEO Excel\testFiles\*.csv')
keyword = "ChartEvents"
keyword2 = "prescriptions"
keyword3 = "Patients"
for file in files:
    if keyword in file:
        df = pd.read_csv(file, sep= '\t',
                    names= ['ROW_ID', 'SUBJECT_ID', 'HADM_ID', 'ICUSTAY_ID', 'ITEMID', 'CHARTTIME', 'STORETIME', 'CGID', 'VALUE', 'VALUENUM', 'VALUEUOM', 'WARNING', 'ERROR', 'RESULTSTATUS'])
    elif keyword2 in file:
        af = pd.read_csv(file, sep= '\t',
                    names= ['Prescription_ID', 'ROW_ID', 'SUBJECT_ID', 'HADM_ID', 'ICUSTAY_ID', 'STARTDATE', 'ENDDATE', 'DRUG_TYPE', 'DRUG', 'DRUG_NAME_POE', 'DRUG_NAME_GENERIC', 'FORMULARY_DRUG_CD', 'GSN', 'NDC', 'PROD_STRENGTH', 'DOSE_VAL_RX', 'DOSE_UNIT_RX', 'FORM_VAL_DISP', 'FORM_UNIT_DISP', 'ROUTE'])
    elif keyword3 in file:
        pf = pd.read_csv(file, sep = '\t',
                    names= ['ROW_ID', 'SUBJECT_ID', 'GENDER', 'DOB', 'DOD', 'DOD_HOSP', 'DOD_SSN', 'EXPIRE_FLAG', 'PRESCRIPTION_IDS', 'PRESCRIPTIONS', 'LABEVENT_IDS', 'LABEVENTS', 'MICROBIOLOGYEVENTS_IDS', 'MICROBIOLOGYEVENTS', 'HADM_IDS', 'HADMS'])

    writer = pd.ExcelWriter(file[:-4] + '.xlsx', engine='xlsxwriter')
    
    workbook = writer.book


    #Heart Rate
    HR = df[(df['ITEMID'] == 220045) & (df['ITEMID'].notnull())]
    HR['CHARTTIME'] = pd.to_datetime(HR['CHARTTIME'])
    HRS = HR[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
    HRS.rename(columns= {'VALUENUM' : 'Heart Rate'}, inplace=True)
    #print(HRS)

    #Systolic Blood Pressure
    BPS = df[(df['ITEMID'] == 220179) & (df['ITEMID'].notnull())]
    BPS['CHARTTIME'] = pd.to_datetime(BPS['CHARTTIME'])
    BPSS = BPS[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
    BPSS.rename(columns= {'VALUENUM' : 'Systolic BP'}, inplace=True)

    #Diatolic Blood Pressure
    BPD = df[(df['ITEMID'] == 220180) & (df['ITEMID'].notnull())]
    BPD['CHARTTIME'] = pd.to_datetime(BPD['CHARTTIME'])
    BPDS = BPD[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
    BPDS.rename(columns= {'VALUENUM' : 'Diatolic BP'}, inplace=True)

    #Blood Pressure Totals
    BPT = pd.merge(BPDS, BPSS, on='CHARTTIME', how='outer')
    #print(BPT)

    #Respiratory Rate
    RR = df[(df['ITEMID'] == 220210) & (df['ITEMID'].notnull())]
    RR['CHARTTIME'] = pd.to_datetime(RR['CHARTTIME'])
    RRS = RR[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
    RRS.rename(columns= {'VALUENUM' : 'Respiratory Rate'}, inplace=True)
    #print(RRS)

    #O2 Saturation
    O2 = df[(df['ITEMID'] == 220277) & (df['ITEMID'].notnull())]
    O2['CHARTTIME'] = pd.to_datetime(O2['CHARTTIME'])
    O2S = O2[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
    O2S.rename(columns= {'VALUENUM' : 'O2 Saturation'}, inplace=True)
    #print(O2S)

    #Temperature
    TP = df[(df['ITEMID'] == 223761) & (df['ITEMID'].notnull())]
    TP['CHARTTIME'] = pd.to_datetime(TP['CHARTTIME'])
    TPS = TP[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
    TPS.rename(columns= {'VALUENUM' : 'Temperature'}, inplace=True)

    #Complete Vitals Chart
    total_Vitals = pd.merge(HRS, BPSS, how='left', on=['CHARTTIME'])
    total_Vitals2 = pd.merge(total_Vitals, BPDS, how='left', on=['CHARTTIME'])
    total_Vitals3 = pd.merge(total_Vitals2, RRS, how='left', on=['CHARTTIME'])
    total_Vitals4 = pd.merge(total_Vitals3, O2S, how='left', on=['CHARTTIME'])
    total_Vitals5 = pd.merge(total_Vitals4, TPS, how='left', on=['CHARTTIME'])
    #print(total_Vitals5)
    total_Vitals5.to_excel( writer, sheet_name='Visualization')
    total_Vitals5.CHARTTIME = pd.to_datetime(df.CHARTTIME)
    total_Vitals5.set_index('CHARTTIME', inplace=True)

    # Create a chart object.
    chart = workbook.add_chart({"type": "line", 'width': 0.25})

    # Configure the series of the chart from the dataframe data.
    # [sheetname, first_row, first_col, last_row, last_col]
    row = len(total_Vitals5.index)
    chart.add_series({
            'name':       [ "Visualization", 0, 2],
            'categories': [ "Visualization", 1, 1, row, 1],
            'values':     [ "Visualization", 1, 2, row, 2],
            'marker':     { 'type': 'circle' },
            'line':       { 'width': 0.25 }
            })

    chart.add_series({
            'name':       [ "Visualization", 0, 3],
            'categories': [ "Visualization", 1, 1, row, 1],
            'values':     [ "Visualization", 1, 3, row, 3],
            'marker':     { 'type': 'circle' },
            'line':       { 'width': 0.25 }
            })

    chart.add_series({
            'name':       [ "Visualization", 0, 4],
            'categories': [ "Visualization", 1, 1, row, 1],
            'values':     [ "Visualization", 1, 4, row, 4],
            'marker':     { 'type': 'circle' },
            'line':       { 'width': 0.25 }
            })
            
    chart.add_series({
            'name':       [ "Visualization", 0, 5],
            'categories': [ "Visualization", 1, 1, row, 1],
            'values':     [ "Visualization", 1, 5, row, 5],
            'marker':     { 'type': 'circle' },
            'line':       { 'width': 0.25 }
            })

    chart.add_series({
            'name':       [ "Visualization", 0, 6],
            'categories': [ "Visualization", 1, 1, row, 1],
            'values':     [ "Visualization", 1, 6, row, 6],
            'marker':     { 'type': 'circle' },
            'line':       { 'width': 0.25 }
            })

    chart.add_series({
            'name':       [ "Visualization", 0, 7],
            'categories': [ "Visualization", 1, 1, row, 1],
            'values':     [ "Visualization", 1, 7, row, 7],
            'marker':     { 'type': 'circle' },
            'line':       { 'width': 0.25 }
            })
            
    chart.set_title({"name": '4 Vitals Visualization'})
    chart.set_x_axis({'text_axis': True, 'name': 'Date'})
    chart.show_blanks_as('span')

    #total_Vitals5.plot()
    #plt.savefig('python.png')

    #Add Medications & GCS
    workbook = writer.book
    worksheet = workbook.add_worksheet('Report')
    writer.sheets['Report'] = worksheet
    header = "Code Status"
    worksheet.write_string(0, 0, header)

    GCS_Verbal = df[(df['ITEMID'] == 223900) & (df['ITEMID'].notnull())]
    GCS_Verbal['CHARTTIME'] = pd.to_datetime(GCS_Verbal['CHARTTIME'])
    GCS_Verbals = GCS_Verbal[["CHARTTIME", "VALUE"]].sort_values(by="CHARTTIME")
    GCS_Verbals.CHARTTIME = pd.to_datetime(df.CHARTTIME)
    GCS_Verbals.rename(columns= {'VALUE' : 'GCS: Verbal'}, inplace=True)
    chart1 = GCS_Verbals.set_index('CHARTTIME').T
    chart1.to_excel(writer, sheet_name ='Report',startrow = 1 , startcol = 0)

    GCS_Motor = df[(df['ITEMID'] == 223901) & (df['ITEMID'].notnull())]
    GCS_Motor['CHARTTIME'] = pd.to_datetime(GCS_Motor['CHARTTIME'])
    GCS_Motors = GCS_Motor[["CHARTTIME", "VALUE"]].sort_values(by="CHARTTIME")
    GCS_Motors.CHARTTIME = pd.to_datetime(df.CHARTTIME)
    GCS_Motors.rename(columns= {'VALUE' : 'GCS: Motor'}, inplace=True)
    chart2 = GCS_Motors.set_index('CHARTTIME').T
    chart2.to_excel(writer, sheet_name ='Report',startrow = chart1.shape[0] + 3, startcol = 0)

    worksheet.write_string(0, 1, "Full Code")
    worksheet.insert_chart('B8', chart)

    #Setup
    writer.save()
    workbook.close