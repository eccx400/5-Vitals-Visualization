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

f = open(r"C:\Users\datiphy\Documents\NEO Excel\Charts\subjects.txt", "r")

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

        #Converts Celcius to Fahrenheit
        def f(x):
                x = x * 1.8 + 32
                return float(x)

        #Calculates 5 Vitals
        def vitals(item_1, item_2, vital_name):
                dataframe = df[(df['ITEMID'] == item_1) | (df['ITEMID'] == item_2)]
                dataframe['CHARTTIME'] = pd.to_datetime(dataframe['CHARTTIME']).dt.strftime('%m-%d, %H:%M')
                dataframes = dataframe[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
                dataframes.rename(columns= {'VALUENUM' : vital_name}, inplace=True)   
                return dataframes

        #Calculates Alarm values
        def alarms(dataframe, alarm_1, alarm_2, vital_name, alarm_name):
                if(alarm_1 == None):
                        alarm = dataframe[(dataframe[vital_name] <= alarm_2)]
                        alarm.rename(columns= {vital_name : alarm_name}, inplace=True)
                elif(alarm_2 == None):
                        alarm = dataframe[(dataframe[vital_name] >= alarm_1)]
                        alarm.rename(columns= {vital_name : alarm_name}, inplace=True)
                else:
                        alarm = dataframe[(dataframe[vital_name] >= alarm_1) | (dataframe[vital_name] <= alarm_2)]
                        alarm.rename(columns= {vital_name : alarm_name}, inplace=True)
                return alarm

        #Calculates GCS Vitals
        def gcs_vitals(item_1, vital_name):
                dataframe = df[(df['ITEMID'] == item_1)]
                dataframes = dataframe[["CHARTTIME", "VALUE"]].sort_values(by="CHARTTIME")
                dataframes.CHARTTIME = pd.to_datetime(dataframe.CHARTTIME)
                dataframes.rename(columns= {'VALUE' : vital_name}, inplace=True)
                return dataframes

        #Gets prescription columns
        def get_prescriptions(pres_name):
                prescription = pres[pres_name]
                p_header = prescription["DRUG"].values[0]
                prescriptions = prescription[["STARTDATE", "DOSE_VAL_RX", "FORM_UNIT_DISP"]].sort_values(by="STARTDATE")
                prescriptions["DOSE_VAL_RX"] = prescriptions["DOSE_VAL_RX"] +" "+ prescriptions["FORM_UNIT_DISP"]
                prescriptions.STARTDATE = pd.to_datetime(prescriptions.STARTDATE)
                prescriptions = prescriptions.drop(columns=["FORM_UNIT_DISP"])
                prescriptions.rename(columns= {'STARTDATE': 'CHARTTIME', 'DOSE_VAL_RX' : p_header}, inplace=True)
                return prescriptions

        #Heart Rate
        HRS = vitals(220045, 211, 'Heart Rate')

        #Heart Rate Alarm
        HR_alarm = alarms(HRS, 140, 30, 'Heart Rate', 'HR Alarm')

        #Systolic Blood Pressure
        BPSS = vitals(220179, 455, 'Systolic BP')

        #Blood Pressure Alarm
        BPS_alarm = alarms(BPSS, None, 80, 'Systolic BP', 'BP Alarm')

        #Diastolic Blood Pressure
        BPDS = vitals(220180, 8441, 'Diastolic BP')

        #Blood Pressure Totals
        BPT = pd.merge(BPSS, BPDS, on='CHARTTIME', how='left')

        #Respiratory Rate
        RRS = vitals(220210, 618, "Respiratory Rate")

        #Respiratory Rate Alarm
        RR_alarm = alarms(RRS, 37, 4, 'Respiratory Rate', 'RR Alarm')

        #O2 Saturation
        O2S = vitals(220277, 646, 'O2 Saturation')

        #Temperature
        TPF = df[(df['ITEMID'] == 223761) | (df['ITEMID'] == 678 )]
        TPC = df[(df['ITEMID'] == 223762) | (df['ITEMID'] == 676 )]
        TPC['ITEMID'] = TPC['ITEMID'].apply(f)
        TP = pd.concat([TPF, TPC])
        TP['CHARTTIME'] = pd.to_datetime(TP['CHARTTIME']).dt.strftime('%m-%d, %H:%M')
        TPS = TP[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
        TPS.rename(columns= {'VALUENUM' : 'Temperature'}, inplace=True)

        #GCS_Verbal
        GCS_Verbals = gcs_vitals(223900, 'GCS: Verbal')

        #GCS_Motor
        GCS_Motors = gcs_vitals(223901, 'GCS: Motor')

        #GCS_Total
        GCS_Totals = gcs_vitals(198, 'GCS: Total')

        try:
                #Gets the top 3 most common prescriptions
                n = 3
                lister = af["GSN"].value_counts().nlargest(3).index.tolist()

                sodium = af[(af['GSN'] == lister[0]) & (af['GSN'].notnull())]
                sodiums = sodium[["STARTDATE", "DOSE_VAL_RX", "FORM_UNIT_DISP"]].sort_values(by="STARTDATE")
                s_header = sodium["DRUG"].values[0]
                sodiums["DOSE_VAL_RX"] = sodiums["DOSE_VAL_RX"] +" "+ sodiums["FORM_UNIT_DISP"]
                sodiums.STARTDATE = pd.to_datetime(sodiums.STARTDATE)
                sodiums = sodiums.drop(columns=["FORM_UNIT_DISP"])
                sodiums.rename(columns= {'STARTDATE': 'CHARTTIME', 'DOSE_VAL_RX' : s_header}, inplace=True)

                fur = af[(af['GSN'] == lister[1]) & (af['GSN'].notnull())]
                furs = fur[["STARTDATE", "DOSE_VAL_RX", "FORM_UNIT_DISP"]].sort_values(by="STARTDATE")
                f_header = fur["DRUG"].values[0]
                furs["DOSE_VAL_RX"] = furs["DOSE_VAL_RX"] +" "+ furs["FORM_UNIT_DISP"]
                furs.STARTDATE = pd.to_datetime(furs.STARTDATE)
                furs = furs.drop(columns=["FORM_UNIT_DISP"])
                furs.rename(columns= {'STARTDATE': 'CHARTTIME', 'DOSE_VAL_RX' : f_header}, inplace=True)

                pro = af[(af['GSN'] == lister[2]) & (af['GSN'].notnull())]
                pros = pro[["STARTDATE", "DOSE_VAL_RX", "FORM_UNIT_DISP"]].sort_values(by="STARTDATE")
                p_header = pro["DRUG"].values[0]
                pros["DOSE_VAL_RX"] = pros["DOSE_VAL_RX"] +" "+ pros["FORM_UNIT_DISP"]
                pros.STARTDATE = pd.to_datetime(pros.STARTDATE)
                pros = pros.drop(columns=["FORM_UNIT_DISP"])
                pros.rename(columns= {'STARTDATE': 'CHARTTIME', 'DOSE_VAL_RX' : p_header}, inplace=True)
        except:
                continue

        #Complete Vitals Chart
        total_Vitals = pd.merge(HRS, BPSS, how='left', on=['CHARTTIME'])
        total_Vitals2 = pd.merge(total_Vitals, BPDS, how='left', on=['CHARTTIME'])
        total_Vitals3 = pd.merge(total_Vitals2, RRS, how='left', on=['CHARTTIME'])
        total_Vitals4 = pd.merge(total_Vitals3, O2S, how='left', on=['CHARTTIME'])
        total_Vitals5 = pd.merge(total_Vitals4, TPS, how='left', on=['CHARTTIME'])
        total_Vitals6 = pd.merge(total_Vitals5, HR_alarm, how='left', on=['CHARTTIME'])
        total_Vitals7 = pd.merge(total_Vitals6, BPS_alarm, how='left', on=['CHARTTIME'])
        total_Vitals8 = pd.merge(total_Vitals7, RR_alarm, how='left', on=['CHARTTIME'])
        total_Vitals8.to_excel( writer, sheet_name='Visualization')

        GCS_Vitals = pd.merge(GCS_Verbals, GCS_Motors, how='outer', on=['CHARTTIME'])
        GCS_Vitals2 = pd.merge(GCS_Vitals, GCS_Totals, how='outer', on=['CHARTTIME'])
        GCS_Vitals3 = pd.merge(GCS_Vitals2, sodiums, how='outer', on=['CHARTTIME'])
        GCS_Vitals4 = pd.merge(GCS_Vitals3, furs, how='outer', on=['CHARTTIME'])
        GCS_Vitals5 = pd.merge(GCS_Vitals4, pros, how='outer', on=['CHARTTIME'])
        GCS_Vitals5 = GCS_Vitals5.sort_values(by="CHARTTIME")
        GCS_Vitals5['CHARTDATE'] = GCS_Vitals5['CHARTTIME']
        GCS_Vitals5.CHARTDATE = pd.to_datetime(GCS_Vitals5.CHARTTIME).dt.strftime('%m-%d')
        GCS_Vitals5.CHARTTIME = pd.to_datetime(GCS_Vitals5.CHARTTIME).dt.strftime('%H:%M')
        cols_to_move = ['CHARTDATE', 'CHARTTIME', 'GCS: Verbal', 'GCS: Motor', 'GCS: Total']
        GCS_Vitals5 =  GCS_Vitals5[ cols_to_move + [ col for col in  GCS_Vitals5.columns if col not in cols_to_move ] ]

        # Create a chart object.
        chart = workbook.add_chart({"type": "line"})

        # Configure the series of the chart from the dataframe data.
        # [sheetname, first_row, first_col, last_row, last_col]
        row = len(total_Vitals8.index)
        
        #HR
        chart.add_series({
                'name':       [ "Visualization", 0, 2],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 2, row, 2],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': '#f15854'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': '#f15854'}
                })

        #BPS
        chart.add_series({
                'name':       [ "Visualization", 0, 3],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 3, row, 3],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': '#faa43a'}, 'border': {'color': 'black'}},
                'line':       { 'width': 1, 'color': '#faa43a'}
                })

        #BPD
        chart.add_series({
                'name':       [ "Visualization", 0, 4],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 4, row, 4],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': '#60bd68'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': '#60bd68'}
                })
                
        #RR
        chart.add_series({
                'name':       [ "Visualization", 0, 5],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 5, row, 5],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': '#5da5da'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': '#5da5da'}
                })

        #O2
        chart.add_series({
                'name':       [ "Visualization", 0, 6],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 6, row, 6],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': '#b276b2'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': '#b276b2'}
                })

        #TP
        chart.add_series({
                'name':       [ "Visualization", 0, 7],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 7, row, 7],
                'marker':     { 'type': 'circle', 'size': 4, 'fill': {'color': '#868686'}, 'border': {'color': 'black'} },
                'line':       { 'width': 1, 'color': '#868686'}
                })

        #HR_alarm
        chart.add_series({
                'name':       [ "Visualization", 0, 8],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 8, row, 8],
                'marker':     { 'type': 'diamond', 'size': 8, 'fill': {'color': 'red'}, 'border': {'color': 'black'} },
                'line':       { 'none': 1 },
                'data_labels':{ 'value': 1, 'position': 'top', 'font': {'size' : 11, 'bold': 1, 'color': 'red'}}
                })

        #BP_alarm
        chart.add_series({
                'name':       [ "Visualization", 0, 9],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 9, row, 9],
                'marker':     { 'type': 'diamond', 'size': 8, 'fill': {'color': 'red'}, 'border': {'color': 'black'} },
                'line':       { 'none': 1 },
                'data_labels':{ 'value': 1, 'position': 'top', 'font': {'size' : 11, 'bold': 1, 'color': 'red'}}
                })

        #RR_alarm
        chart.add_series({
                'name':       [ "Visualization", 0, 10],
                'categories': [ "Visualization", 1, 1, row, 1],
                'values':     [ "Visualization", 1, 10, row, 10],
                'marker':     { 'type': 'diamond', 'size': 8, 'fill': {'color': 'red'}, 'border': {'color': 'black'} },
                'line':       { 'none': 1 },
                'data_labels':{ 'value': 1, 'position': 'top', 'font': {'size' : 11, 'bold': 1, 'color': 'red'}}
        })
                
        chart.set_title({"name": '5 Vitals Visualization'})
        chart.set_x_axis({'text_axis': True, 'name': 'Date'})
        chart.set_legend({"none": 1})
        chart.show_blanks_as('span')

        #Add Medications & GCS
        workbook = writer.book
        worksheet = workbook.add_worksheet('Report')
        writer.sheets['Report'] = worksheet
        header = "Code Status"

        worksheet.insert_chart('A1', chart, {'x_scale': 7.777, 'y_scale': 1.944})

        worksheet.write_string(29, 0, header)
        worksheet.write_string(29, 1, "Full Code")

        chart1 = GCS_Vitals5.set_index('CHARTDATE').T
        chart1.to_excel(writer, sheet_name ='Report', startrow = 30 , startcol = 0)
        
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