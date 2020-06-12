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

__author__ = "Eric Cheng"

f = open(r"C:\Users\datiphy\Documents\NEO Excel\Charts\icd_subjects.txt", "r")

newpath = "C:\\Users\\datiphy\\Documents\\NEO Excel\\icd_files\\"
txtlines = f.readlines()
for line in txtlines:
        subject_id = line.strip()
        chartfile= "C:\\Users\\datiphy\\Documents\\NEO Excel\\46520\\" + "subject_"+ subject_id + "_ChartEvents.txt"
        prescfile = "C:\\Users\\datiphy\\Documents\\NEO Excel\\46520_P\\" + "subject_" + subject_id + "_prescriptions.csv"

        #CHART EVENTS
        df = pd.read_csv(chartfile, sep= '\t', index_col= False, names= ['ROW_ID', 'SUBJECT_ID', 'HADM_ID', 'ICUSTAY_ID', 'ITEMID', 'CHARTTIME', 'STORETIME', 'CGID', 'VALUE', 'VALUENUM', 'VALUEUOM', 'WARNING', 'ERROR', 'RESULTSTATUS']) 

        #PRESCRIPTION
        af = pd.read_csv(prescfile, sep= '\t', names= ['Prescription_ID', 'ROW_ID', 'SUBJECT_ID', 'HADM_ID', 'ICUSTAY_ID', 'STARTDATE', 'ENDDATE', 'DRUG_TYPE', 'DRUG', 'DRUG_NAME_POE', 'DRUG_NAME_GENERIC', 'FORMULARY_DRUG_CD', 'GSN', 'NDC', 'PROD_STRENGTH', 'DOSE_VAL_RX', 'DOSE_UNIT_RX', 'FORM_VAL_DISP', 'FORM_UNIT_DISP', 'ROUTE'])
        print(subject_id)

        writer = pd.ExcelWriter(newpath + subject_id + '.xlsx', engine='xlsxwriter')
        workbook = writer.book

        #Heart Rate
        HR = df[(df['ITEMID'] == 220045) | (df['ITEMID'] == 211)]
        HR['CHARTTIME'] = pd.to_datetime(HR['CHARTTIME']).dt.strftime('%m-%d, %H:%M')
        HRS = HR[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
        HRS.rename(columns= {'VALUENUM' : 'Heart Rate'}, inplace=True)

        #Systolic Blood Pressure
        BPS = df[(df['ITEMID'] == 220179) | (df['ITEMID'] == 455)]
        BPS['CHARTTIME'] = pd.to_datetime(BPS['CHARTTIME']).dt.strftime('%m-%d, %H:%M')
        BPSS = BPS[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
        BPSS.rename(columns= {'VALUENUM' : 'Systolic BP'}, inplace=True)

        #Diatolic Blood Pressure
        BPD = df[(df['ITEMID'] == 220180) | (df['ITEMID'] == 8441)]
        BPD['CHARTTIME'] = pd.to_datetime(BPD['CHARTTIME']).dt.strftime('%m-%d, %H:%M')
        BPDS = BPD[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
        BPDS.rename(columns= {'VALUENUM' : 'Diatolic BP'}, inplace=True)

        #Blood Pressure Totals
        BPT = pd.merge(BPDS, BPSS, on='CHARTTIME', how='outer')

        #Respiratory Rate
        RR = df[(df['ITEMID'] == 220210) | (df['ITEMID'] == 618)]
        RR['CHARTTIME'] = pd.to_datetime(RR['CHARTTIME']).dt.strftime('%m-%d, %H:%M')
        RRS = RR[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
        RRS.rename(columns= {'VALUENUM' : 'Respiratory Rate'}, inplace=True)

        #O2 Saturation
        O2 = df[(df['ITEMID'] == 220277) | (df['ITEMID'] == 646)]
        O2['CHARTTIME'] = pd.to_datetime(O2['CHARTTIME']).dt.strftime('%m-%d, %H:%M')
        O2S = O2[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
        O2S.rename(columns= {'VALUENUM' : 'O2 Saturation'}, inplace=True)

        #Temperature
        TP = df[(df['ITEMID'] == 223761) | (df['ITEMID'] == 678 )]
        TP['CHARTTIME'] = pd.to_datetime(TP['CHARTTIME']).dt.strftime('%m-%d, %H:%M')
        TPS = TP[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
        TPS.rename(columns= {'VALUENUM' : 'Temperature'}, inplace=True)

        #GCS_Verbal
        GCS_Verbal = df[(df['ITEMID'] == 223900)]
        GCS_Verbals = GCS_Verbal[["CHARTTIME", "VALUE"]].sort_values(by="CHARTTIME")
        GCS_Verbals.CHARTTIME = pd.to_datetime(GCS_Verbals.CHARTTIME).dt.strftime('%m-%d, %H:%M')
        GCS_Verbals.rename(columns= {'VALUE' : 'GCS: Verbal'}, inplace=True)
        print(GCS_Verbals)

        #GCS_Motor
        GCS_Motor = df[(df['ITEMID'] == 223901)]
        GCS_Motors = GCS_Motor[["CHARTTIME", "VALUE"]].sort_values(by="CHARTTIME")
        GCS_Motors.CHARTTIME = pd.to_datetime(GCS_Motors.CHARTTIME).dt.strftime('%m-%d, %H:%M')
        GCS_Motors.rename(columns= {'VALUE' : 'GCS: Motor'}, inplace=True)
        print(GCS_Motors)

        #GCS_Total
        GCS_Total = df[(df['ITEMID'] == 198)]
        GCS_Totals = GCS_Total[["CHARTTIME", "VALUE"]].sort_values(by="CHARTTIME")
        GCS_Totals.CHARTTIME = pd.to_datetime(GCS_Totals.CHARTTIME).dt.strftime('%m-%d, %H:%M')
        GCS_Totals.rename(columns= {'VALUE' : 'GCS: Total'}, inplace=True)
        print(GCS_Totals)

        #Complete Vitals Chart
        total_Vitals = pd.merge(HRS, BPSS, how='left', on=['CHARTTIME'])
        total_Vitals2 = pd.merge(total_Vitals, BPDS, how='left', on=['CHARTTIME'])
        total_Vitals3 = pd.merge(total_Vitals2, RRS, how='left', on=['CHARTTIME'])
        total_Vitals4 = pd.merge(total_Vitals3, O2S, how='left', on=['CHARTTIME'])
        total_Vitals5 = pd.merge(total_Vitals4, TPS, how='left', on=['CHARTTIME'])
        GCS_Vitals = pd.merge(GCS_Verbals, GCS_Motors, how='outer', on=['CHARTTIME'])
        GCS_Vitals2 = pd.merge(GCS_Vitals, GCS_Totals, how='outer', on=['CHARTTIME'])
        GCS_Vitals2 = GCS_Vitals2[['CHARTTIME', 'GCS: Verbal', 'GCS: Motor', 'GCS: Total']]
        total_Vitals5.to_excel( writer, sheet_name='Visualization')

        # Create a chart object.
        chart = workbook.add_chart({"type": "line"})

        # Configure the series of the chart from the dataframe data.
        # [sheetname, first_row, first_col, last_row, last_col]
        row = len(total_Vitals5.index)
        chart.add_series({
                'name':       [ "Visualization", 0, 2],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 2, row, 2],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': 'black'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': 'black'}
                })

        chart.add_series({
                'name':       [ "Visualization", 0, 3],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 3, row, 3],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': 'red'}, 'border': {'color': 'black'}},
                'line':       { 'width': 1, 'color': 'red'}
                })

        chart.add_series({
                'name':       [ "Visualization", 0, 4],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 4, row, 4],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': 'blue'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': 'blue'}
                })
                
        chart.add_series({
                'name':       [ "Visualization", 0, 5],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 5, row, 5],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': 'orange'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': 'orange'}
                })

        chart.add_series({
                'name':       [ "Visualization", 0, 6],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 6, row, 6],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': 'purple'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': 'purple'}
                })

        chart.add_series({
                'name':       [ "Visualization", 0, 7],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 7, row, 7],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': 'yellow'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': 'yellow'}
                })
                
        chart.set_title({"name": '4 Vitals Visualization'})
        chart.set_x_axis({'text_axis': True, 'name': 'Date'})
        chart.set_legend({"none": 1})
        chart.show_blanks_as('span')

        #Add Medications & GCS
        worksheet = workbook.add_worksheet('Report')
        writer.sheets['Report'] = worksheet
        header = "Code Status"

        worksheet.insert_chart('A1', chart, {'x_scale': 7.777, 'y_scale': 1.944})

        worksheet.write_string(29, 0, header)
        worksheet.write_string(29, 1, "Full Code")

        try:                                
                chart1 = GCS_Vitals2.set_index('CHARTTIME').T
                chart1.to_excel(writer, sheet_name ='Report', startrow = 30 , startcol = 0)

                n = 3
                lister = af["GSN"].value_counts().nlargest(3).index.tolist()

                sodium = af[(af['GSN'] == lister[0]) & (af['GSN'].notnull())]
                sodiums = sodium[["DRUG", "DOSE_VAL_RX"]].sort_values(by="DRUG")
                chart4 = sodiums.set_index('DRUG').T
                chart4.to_excel(writer, sheet_name ='Report',startrow = 35 , startcol = 0)

                fur = af[(af['GSN'] == lister[1]) & (af['GSN'].notnull())]
                furs = fur[["DRUG", "DOSE_VAL_RX"]].sort_values(by="DRUG")
                chart5 = furs.set_index('DRUG').T
                chart5.to_excel(writer, sheet_name ='Report',startrow = 38, startcol = 0)

                pro = af[(af['GSN'] == lister[2]) & (af['GSN'].notnull())]
                pros = pro[["DRUG", "DOSE_VAL_RX"]].sort_values(by="DRUG")
                chart6 = pros.set_index('DRUG').T
                chart6.to_excel(writer, sheet_name ='Report',startrow = 41, startcol = 0)
        except:
                continue
        
        #Setup
        writer.save()
        workbook.close

path1 = 'C:\\Users\\datiphy\\Documents\\NEO Excel\\Charts\\ADDSv3.xlsm'
xl = Dispatch("Excel.Application")
wb1 = xl.Workbooks.Open(path1)
for filename in glob.glob('C:\\Users\\datiphy\\Documents\\NEO Excel\\icd_files\\*.xlsx'):
        print(filename)
        try:
                wb2 = xl.Workbooks.Open(filename)
                ws1 = wb1.Worksheets(1)
                ws1.Copy(Before=wb2.Worksheets(1))
                wb2.Close(SaveChanges=True)
        except:
                print
xl.Quit()
