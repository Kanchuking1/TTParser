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

outdf = pd.DataFrame([],columns=["Slot", "CourseCap"])

ttslot = OrderedDict()
ttcap = OrderedDict()

for a in ttparse["COURSENO"].tolist():
    #print(a)
    temp = ttparse[ttparse["COURSENO"] == a][ttparse["STAT"] == 'L']["DAYS/ H"].tolist()
    #REMOVE DUPLICATES FROM TEMP
    temp = list(OrderedDict.fromkeys(temp))
    ttslot[a] = temp[0]
    cap = 0
    if (not ttcap.__contains__(a)):
        for i in ttparse[ttparse["COURSENO"] == a]["CAPACITY"].tolist():
            if(not math.isnan(i)):
                cap = cap + int(i)
    
        ttcap[a] = cap

#Making the new format
ttcourse = OrderedDict()

for a in ttslot:
    ttcourse[ttslot[a]] = []

for a in ttslot:
    temp = []
    temp.append(a)
    temp.append(ttcap[a])
    ttcourse[ttslot[a]].append(temp)

for a in ttcourse:
    outdf = outdf.append({
        'Slot': a,
        'CourseCap': ttcourse[a]
    },
    ignore_index= True)

print(outdf)

outdf.to_json("tt_parsed_slottocourse.json", orient='table')