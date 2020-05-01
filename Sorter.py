import json
import xlsxwriter
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

df = pd.read_csv (r'C:\Users\datiphy\Documents\NEO Excel\Charts\65278.csv')

print("Column headings:")
print(df.columns)

writer = pd.ExcelWriter('65278_Report.xlsx', engine='xlsxwriter')

HR = df[(df['ITEMID'] == 220045) & (df['ITEMID'].notnull())]
HRS = HR[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
#plt.plot(HRS['CHARTTIME'], HRS['VALUENUM']) #For multiple plots
HRS.plot(x ='CHARTTIME', y='VALUENUM', kind= "line")
print(HRS)
HRS.to_excel( writer, sheet_name='Heart Rate')

#Do something about Blood Pressure
#BPS = df[(df['ITEMID'] == 220179) & (df['ITEMID'].notnull())]
#BPSS = BPS.sort_values(by="CHARTTIME")
#print(BPSS)

#BPD = df[(df['ITEMID'] == 220180) & (df['ITEMID'].notnull())]
#BPDS = BPD.sort_values(by="CHARTTIME")
#print(BPDS)

RR = df[(df['ITEMID'] == 220210) & (df['ITEMID'].notnull())]
RRS = RR[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
RRS.plot(x ='CHARTTIME', y='VALUENUM', kind= "line")
print(RRS)
RRS.to_excel( writer, sheet_name='Respiratory Rate')

O2 = df[(df['ITEMID'] == 220277) & (df['ITEMID'].notnull())]
O2S = O2[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
O2S.plot(x ='CHARTTIME', y='VALUENUM', kind= "line")
print(O2S)
O2S.to_excel( writer, sheet_name='O2 Saturation')

TP = df[(df['ITEMID'] == 223761) & (df['ITEMID'].notnull())]
TPS = TP[["CHARTTIME", "VALUENUM"]].sort_values(by="CHARTTIME")
print(TPS)
TPS.plot(x ='CHARTTIME', y='VALUENUM', kind= "line")
TPS.to_excel( writer, sheet_name='Temperature')

#worksheet.insert_chart('Visualization', chart)
writer.save()
plt.show()