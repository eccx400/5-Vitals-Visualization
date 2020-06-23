import xlsxwriter
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
from functools import reduce
pd.options.mode.chained_assignment = None  # default='warn'

f = open(r"C:\Users\datiphy\Documents\NEO Excel\Charts\updated_rankings.txt", "r")
e = open(r"C:\Users\datiphy\Documents\NEO Excel\Charts\subjects.txt","w")

for i in range(150):
    x = f.readline()
    if x == "": 
        continue
    else:
        x = x.split(" ")[0]
        x = x[0:]
    e.write(x + "\n")

