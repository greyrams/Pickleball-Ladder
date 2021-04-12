# -*- coding: utf-8 -*-
"""
Created on Mon Nov  2 20:14:41 2020

@author: srams
"""

# import packages

import os
import pandas as pd
import numpy as np
from openpyxl import load_workbook

# Set League Name

league = "Winter2021Mon"

# Change working directory

os.chdir("C:\\Users\\srams\OneDrive\Documents\PickleballLadder\Winter2021Mon")
os.listdir('.')

# Get overall results file

initial_file = "WkByWkResults_"+league+".xlsx"
xls_initial = pd.ExcelFile(initial_file)

initial_df = xls_initial.parse('Stats')

WkPos = initial_df["Week"].max()
CurrentWk = int(WkPos + 1)

# Read temp file

FileTemp = "WeekTemp.csv"
CurrentTemp = pd.read_csv(FileTemp)

# Read results file

Results = "Results_"+league+".xlsx"
Resultsxls = pd.ExcelFile(Results)

Resultsdf = Resultsxls.parse("Week"+str(CurrentWk))

Resultsdf = Resultsdf.rename(columns = {'Absent':'AbsentRes','Sub':'SubRes'})

### Read adjustments file - leave for now ###

#AdjGrp = "GrpAdjust.xlsx"
#GrpAdjxls = pd.ExcelFile(AdjGrp)

#GrpAdj = GrpAdjxls.parse(GrpAdjxls.sheet_names[WkPos])

# Merge CurrentTemp with Results

CalcNew = CurrentTemp.merge(Resultsdf, on='Player', how = 'left')

# Calculate max/min scores for each playing group

CalcNew['grpmax'] = CalcNew.groupby(by='WkCourt').Score.transform('max')
CalcNew['grpmin'] = CalcNew.groupby(by='WkCourt').Score.transform('min')

# Set Place for each player

rows = CalcNew.shape[0]

for x in range(0, rows) :
    if CalcNew.loc[x,'Absent'] == "Yes" :
        CalcNew.loc[x,'Place'] = "Absent"
    elif CalcNew.loc[x,'Score'] == CalcNew.loc[x,'grpmax'] :
        CalcNew.loc[x,'Place'] = "Max"
    elif CalcNew.loc[x,'Score'] == CalcNew.loc[x,'grpmin'] :
        CalcNew.loc[x,'Place'] = "Min"
    else :
        CalcNew.loc[x,'Place'] = "Mid"

# See if a sub has min score; if yes then "Yes" in SubMin column for subs in that play group
# and adjust set SubAdj to -0.5

for z in range(0, rows) :
    if CalcNew.loc[z,'Sub'] != "No" and CalcNew.loc[z,'Score'] == CalcNew.loc[z,'grpmin'] :
        CalcNew.loc[z,'SubMin'] = "Yes"
    else :
        CalcNew.loc[z,'SubMin'] = "No"

CalcNew['SubMin'] = CalcNew.groupby(by='WkCourt').SubMin.transform('max')

for y in range(0, rows):
    if CalcNew.loc[y,'SubMin'] == "Yes" :
        CalcNew.loc[y,'SubAdj'] = 0.5
    else :
        CalcNew.loc[y,'SubAdj'] = 0

# Set Base Group

CalcNew['PlayGrpFinal'] = CalcNew['PlayGrp']

CalcNew = CalcNew.rename(columns = {'PlayGrp':'PlayGrpInit'})
CalcNew['PlayGrpInit'] = CalcNew['PlayGrpInit'].astype('object')

for i in range(0, rows) :
    if CalcNew.loc[i,'PlayGrpInit'] != "CT" and CalcNew.loc[i,'Absent'] == "No" :
        CalcNew.loc[i,'BaseGrp'] = (CalcNew.loc[i,'StartGrp']+CalcNew.loc[i,'PlayGrpFinal'])/2
    else :
        CalcNew.loc[i,'BaseGrp'] = CalcNew.loc[i,'StartGrp']+CalcNew.loc[i,'SubAdj']
     
### CalcNew = CalcNew.merge(GrpAdj, on='Player', how = 'left') ###

# Set New Group

for j in range(0, rows) :
    if CalcNew.loc[j,'Place'] == "Min" :
        CalcNew.loc[j, 'NewGrp'] = CalcNew.loc[j, 'BaseGrp']+1
    elif CalcNew.loc[j,'Place'] == "Max" :
        CalcNew.loc[j, 'NewGrp'] = CalcNew.loc[j, 'BaseGrp']-1
    else : 
        CalcNew.loc[j, 'NewGrp'] = CalcNew.loc[j, 'BaseGrp']
        
### CalcNew['NewGrp'] = CalcNew['NewGrp'] + CalcNew['GrpAdj'].fillna(0) ###

# Sort and Set New Rank

CalcNew = CalcNew.sort_values(['NewGrp','StartRank'])
CalcNew = CalcNew.reset_index(drop=True)

CalcNew['EndRank'] = CalcNew.index + 1
CalcNew['EndGrp'] = np.ceil(CalcNew['EndRank'].values/4).astype('int64')

CalcNew['Week'] = CurrentWk

CalcNew = CalcNew.rename(columns = {'PlayGrpFinal':'PlayGrp'})

print(CalcNew[['Player','StartGrp','PlayGrp','BaseGrp','Place','NewGrp','StartRank','EndRank']])
print(CalcNew[['Absent','Score','grpmax','grpmin','Place']])

# Prep Data for Export to Week by Week Resutls

data_to_append = CalcNew[['Player','StartRank','StartGrp','Week','Absent','Sub','PlayGrp','BaseGrp','Score','Place','NewGrp','EndRank','EndGrp']]

DataExport = pd.concat([initial_df, data_to_append])

# Append new data to Week to Week Results

writer = pd.ExcelWriter(initial_file, engine='xlsxwriter')
DataExport.to_excel(writer, sheet_name = 'Stats', index = False)
writer.save()

# Prep attendance for next week

NextWk = CurrentWk + 1
NextAttend = CalcNew[['Player','Absent','Sub']]

NextAttend['Absent'] = "No"
NextAttend['Sub'] = "No"

# Open Attendance file and add list for next week    

Attend = "Attendance_"+league+".xlsx"

writer_att = pd.ExcelWriter(Attend, engine='openpyxl')
writer_att.book = load_workbook(Attend)
writer_att.sheets = dict((wsa.title, wsa) for wsa in writer_att.book.worksheets)

AttendSheet = "Week"+str(NextWk)

try :
    AttSheet = writer_att.book[sheets[AttendSheet]]
    writer_att.sheets[AttendSheet].delete_rows(2,50)
    NextAttend.to_excel(writer, sheet_name = AttendSheet, startrow = 1, index = False, header = False)
except NameError as error :
    NextAttend.to_excel(writer_att, sheet_name = AttendSheet, index = False)
   
writer_att.close()
    


