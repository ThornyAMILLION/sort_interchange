########################################################################### 
# Author: adam
# Date: April 22th, 2022
# File: sort_interchange.py
# Purpose: Sort the interchange file 
###########################################################################

from csv import writer
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import openpyxl

filename = "Interchange Master.xlsx"
filename1 = "nodupe.xlsx"

df1 = pd.read_excel(filename)
df2 = df1[['info', 'productnum']]
df3 = df2

product_dict = {}

for i in df3.values:
    if i[0] not in product_dict.keys():
        product_dict[i[0]] = [i[1]]
    else:
        product_dict[i[0]].append(i[1])

df4 = pd.DataFrame.from_dict(product_dict, orient='index')

writer = ExcelWriter(filename1)
df4.to_excel(writer, 'Sheet1')
writer.save()

print('done')