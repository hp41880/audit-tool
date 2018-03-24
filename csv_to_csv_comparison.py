# -*- coding: utf-8 -*-
"""
Created on Fri Mar 23 18:56:05 2018
This program compares two csv files caf.csv and c1a.csv .Its output template.xlsx has 4 sheets.
cleaned sheets are the one after removing all characters except 0-9 a-z A-Z 
Any field which is in caf but not in c1a will be removed

@author: hp41880@gmail.com
"""

from pandas import read_csv
import sys,csv

#path where csv files reside and tempfile for field mapping should be saved
path="F:\\excel_to_excel_comparison\\"

caf=read_csv(path+"caf.csv")
c1a=read_csv(path+"c1a.csv")

list_of_columns_of_caf=caf.columns.tolist()
list_of_columns_of_c1a=c1a.columns.tolist()

#following snippet will print column names of caf and c1a into temp.csv
rows = zip(list_of_columns_of_caf,list_of_columns_of_c1a)
with open(path+"temp.csv", "w",newline='') as f:
    writer = csv.writer(f)
    for row in rows:
        writer.writerow(row)

print("\n\n")

temp=input("please map second column of temp.csv and save it as fields_map.csv.\n After saving the file, enter 'please continue'\n") 
if temp!='please continue':
    sys.exit("re-run the program and follow above step")
    

fields_map=read_csv(path+"fields_map.csv",header=None)
a=fields_map.set_index(0).to_dict()  #will make dict a[1] will have (keys=columns of caf,vals=columns of c1a)
caf.rename(columns=a[1], inplace=True)  # replace caf column names by c1a column names

#df.rename(columns={'oldName1': 'newName1', 'oldName2': 'newName2'}, inplace=True)

c1a=c1a.astype(str)
caf=caf.astype(str)    
caf_cleaned=caf.replace("[^0-9a-zA-Z]+",'',regex=True)  #removing special chars
c1a_cleaned=c1a.replace("[^0-9a-zA-Z]+",'',regex=True)

#set .name attribute-- it will be our excel sheet name
c1a.name='c1a'
caf.name='caf'
caf_cleaned.name='caf_cleaned'
c1a_cleaned.name='c1a_cleaned'

caf = caf.loc[:, caf.columns.notnull()]    #removing column of name 'nan' ie columns in caf but not in database or caf column we dont want to compare
caf_cleaned = caf_cleaned.loc[:, caf_cleaned.columns.notnull()] #'nan' column needs to be removed so that colnames can be sorted

caf = caf.reindex(columns=sorted(caf.columns))
c1a = c1a.reindex(columns=sorted(c1a.columns))
caf_cleaned = caf_cleaned.reindex(columns=sorted(caf_cleaned.columns))
c1a_cleaned = c1a_cleaned.reindex(columns=sorted(c1a_cleaned.columns))


xlr = pd.ExcelWriter(path+"template.xlsx")
caf.to_excel(xlr, 'caf')
c1a.to_excel(xlr, 'c1a') 
caf_cleaned.to_excel(xlr, 'caf_cleaned')
c1a_cleaned.to_excel(xlr, 'c1a_cleaned')
xlr.save()
         
