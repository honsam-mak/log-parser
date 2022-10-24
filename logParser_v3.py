# Before running this program, need to create an empty excel file 
# named 'tx_result_v3.xslx' with sheet 'all_tx'.
# usage : python3 logParser_v3.py result.xlsx log1 

from datetime import datetime
import pandas as pd
import openpyxl
import sys

if len(sys.argv) <3 :
    print('usage: python3 logParser_v3.py result.xlsx log1')
    sys.exit()

#file name
log_file=sys.argv[2]
xls_file=sys.argv[1]

#Column name
#title_list = ['TX_NAME','PID','START_TM','END_TM','EXE_SEC']
listAll = []
templist = []

count = 0
row = 0 

with open(log_file, encoding="utf-8", errors='ignore') as fp:
    for line in fp:
        count += 1

        #Tx Begin
        if 'Program started...' in line :

            #print out the log
#            print("Line{}: {}".format(count, line.strip()))

            words = line.split()
            index = words.index('Program')
            txname = words[index-2]
            pid = int(words[index+3][5:-1])

            #check if the same pid and txname exists in list
            for element in listAll:
                # check if the pid and tx name are same
                if (element[1] == pid) and (element[0] in txname) :
                    templist.append(element)
                    row += 1
                    listAll.remove(element)
                    break
            
            new_list = [txname,pid,"{} {}".format(words[index-4][-5:],words[index-3]),'','']
            listAll.append(new_list)

        #Tx End
        if 'ComTrxReceive  After ComMQTrxReceive()' in line :

            #print out the log
#            print("Line{}: {}".format(count, line.strip()))

            words = line.split()
            index = words.index('ComTrxReceive')
            fintime = "{} {}".format(words[index-4][-5:],words[index-3])
            txname = words[index-2]
            pid = int(words[index+3][5:-1])

            #update the finish time
            for element in listAll:

                # check if the pid and tx name are same
                if (element[1] == pid) and (element[0] in txname) :
                    element[3] = fintime
                    dt1 = datetime.strptime(element[2],"%m/%d %H:%M:%S")
                    dt2 = datetime.strptime(element[3],"%m/%d %H:%M:%S")
                    delta = dt2 - dt1
                    element[4] = delta.total_seconds()

res_list = listAll + templist

wb = openpyxl.load_workbook(xls_file) 
ws = wb.active
for tx in res_list:
    ws.append(tx)

wb.save(xls_file)

