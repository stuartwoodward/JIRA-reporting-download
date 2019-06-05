 #Use "pip install pandas" in Python from command prompt if you dont have pandas already installed
#Use "pip install openpyxl" in Python from command prompt to install openpyx1 required to export to excel
import pandas as pd
import os
import csv


Column_Filter = []
li = []
HF = 'C:\\Users\\stuart.woodward\\Desktop\\06. Python\\Config\\HeaderFilter.csv'
outputfile = 'C:\\Users\\stuart.woodward\\Desktop\\06. Python\\All projects Intelehub (Jira).xlsx'
excel_sheet = 'All projects Intelehub (Jira)'


#Get the columns to import from a csv file
with open(HF, 'r') as f:
  reader = csv.reader(f)
  Column_Filter = list(reader)

#Get the names of all csv files in a directory
all_csv = [file_name for file_name in os.listdir(os.getcwd()) if '.csv' in file_name]

#Import the required columns from multiple csv files and append to 1 file
for filename in all_csv:
    df = pd.read_csv(filename, index_col=None, header=0, parse_dates=True, infer_datetime_format=True,usecols=Column_Filter[0])
    li.append(df)
#print (li)

frame = pd.concat(li, axis=0, ignore_index=True)
frame.to_excel (outputfile, sheet_name = excel_sheet,index = False, header=True)
#frame.to_csv('All projects Intelehub (Jira).csv', index=False)




