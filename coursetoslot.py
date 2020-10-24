#
# 1. Read Input2 TT Sheet name sem 1 19 - 20 TIME TABLE.
# 2. If it is a lecture section then
# 3. Add to capacity to the pre existing dictionary 
# 
# Course format:
# Course ID String
# DAYS/H List
# Capacity Integer
# This script is created by Wahib Sabir Kapdi
# 

import pandas as pd
import math

from collections import OrderedDict

dropProp = ["COM CODE","INSTRUCTOR IN CHARGE/Instructor", "COURSETITLE", "CREDIT                          L P U", "SEC", "ROOM", "COMPRE DATE"]

Unnamed = [10]

for i in range(12,23):
    Unnamed.append(i)

for i in Unnamed:
    dropProp.append("Unnamed: " + str(i))

ttparse = pd.read_excel('Timetable_inp2.xlsx', sheet_name = 'sem 1 19-20 TIME TABLE', header = 2)
#Filtering out all the useless data
ttparse.drop(columns=dropProp, inplace=True)
nan_row = list(ttparse [ttparse["COURSENO"].isnull()].index)
ttparse.drop(nan_row, inplace = True)
ttparse = ttparse[ttparse.STAT != "P"][ttparse.STAT != "R"][ttparse.STAT != "I"]
ttparse[ttparse["DAYS/ H"].isnull()]["DAYS/ H"] = "TBA"
#ttparse.drop(columns = ["STAT"], inplace = True)
print (ttparse)

outdf = pd.DataFrame([],columns=["Course No.", "Slots", "Capacity"])

ttslot = OrderedDict()
ttcap = OrderedDict()

for a in ttparse["COURSENO"].tolist():
    #print(a)
    temp = ttparse[ttparse["COURSENO"] == a][ttparse["STAT"] == 'L']["DAYS/ H"].tolist()
    #REMOVE DUPLICATES FROM TEMP
    temp = list(OrderedDict.fromkeys(temp))
    str1 = ""
    for t in temp:
        str1 = str1 + str(t) + ", "
    if(len(str1) == 0):
        str1 = "NA"
    print (str1)
    ttslot[a] = str1
    cap = 0
    if (not ttcap.__contains__(a)):
        for i in ttparse[ttparse["COURSENO"] == a]["CAPACITY"].tolist():
            if(not math.isnan(i)):
                cap = cap + int(i)
    
        ttcap[a] = cap

#Building the new data frame

for a in ttslot:
    outdf = outdf.append({
        'Course No.': a,
        'Slots' : ttslot[a],
        'Capacity' : ttcap[a]
    },
    ignore_index= True)

outdf.to_excel("tt_parsed.xlsx", index = False)