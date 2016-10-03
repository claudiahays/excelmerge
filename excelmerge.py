
#
# Claudia Bittmann September 2016
#
# Email: coding.claudia@gmail.com
#
# This program pulls in excel files from a specific folder
# and merges them all into a single dataframe and exports the df to a csv
# using python and the below imported modules.
#

from pandas import ExcelWriter
import csv
import pandas
import time
import datetime
import locale
import os
import sys

df_files = pandas.DataFrame()
path = 'C:\\Users\\Claudia\\Documents\\Python\\Test'
for folderName, subfolders, filenames in os.walk(path):
    df_files = df_files.append(filenames)
df_files['Filenames'] = df_files
print(df_files)


x = len(df_files.index)
z = 0

full_df = pandas.DataFrame()

while z < x:
    f = df_files.iloc[z]['Filenames']
    f = str(f)
    if z == 0:

        var = path
        var2 = var + f
        print(var2)
        xl = pandas.ExcelFile(var2)
        df = xl.parse(sheetname=0, skiprows=0)
        full_df = full_df.append(df)
        z = z + 1

    else: 
        var = path
        var2 = var + f
        print(var2)
        xl = pandas.ExcelFile(var2)
        df = xl.parse(sheetname=0, skiprows=0)
        full_df = full_df.append(df)
        z = z + 1

print(full_df)

full_df_2 = full_df.drop_duplicates(['Test 1'], keep='first')
writer = ExcelWriter('test.xlsx')
full_df_2.to_excel(writer,'Sheet1')
writer.save()
