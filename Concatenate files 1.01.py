#Use "pip install pandas" in Python from command prompt if you dont have pandas already installed
#Use "pip install openpyxl" in Python from command prompt to install openpyx1 required to export to excel
import pandas as pd
import os
import csv

li = []
#outputfilepath = 'C:\\Users\\stuart.woodward\\Desktop\\06. Python\\'
outputfilepath = 'C:\\Users\\Stuart.Woodward\\Intelematics Australia Pty Ltd\\DMS - ARC\\02. Continuous improvement\\03. Metrics\\'
#import_field = 'Labels'
import_field = 'Fix versions'
ImportIndex=[]
KeyIndex=[]
#TargetFields =['Fix versions','Labels','Custom field (Services Impacting Deployments)']

outputfile=outputfilepath+import_field+'.xlsx'

#Get the names of all csv files in a directory
all_csv = [file_name for file_name in os.listdir(os.getcwd()) if '.csv' in file_name]

#Import the data from multiple csv files to a panda data frame
for filename in all_csv:
    df = pd.read_csv(filename, index_col=None, header=0, parse_dates=True, infer_datetime_format=True)

    #find the columns that match the Issue key and Import field description
    columns = list(df.columns.values)
    for column in columns:
        if column.startswith(import_field):
           ImportIndex.append(columns.index(column))
        if column.startswith("Issue key"):
           KeyIndex.append(columns.index(column))

    #convert each row of a frame to a tuple
    for row in df.itertuples(index=False):
         
        #create multiple tuples where the import field is repeated
        for column in ImportIndex:
            #Remove rows which are numbers. Nulls are imported from csv as nan in dataframes. 
            #this loop removes the rows that have nan as a value
            if type(row[column]) == str:
                modifiedrows = [str(row[KeyIndex[0]]) ,str(row[column])]
                li.append(modifiedrows)

    #clear the index for each file as the labels can be in different columns
    ImportIndex.clear()
    KeyIndex.clear()

#convert the tuple to a panda data frame
frame = pd.DataFrame(li, columns = ['Issue key',import_field])

#convert the panda data frame to an excel file
frame.to_excel (outputfile, sheet_name = import_field,index = False, header=True)
