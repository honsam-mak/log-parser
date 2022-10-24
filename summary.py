# First run logParser_v3.py
# usage : python3 summary.py result.xlsx

import pandas as pd
import openpyxl
import sys

if len(sys.argv) <2 :
    print('usage: python3 summary.py result.xlsx')
    sys.exit()

#file name
xls_file=sys.argv[1]

excel_data_df = pd.read_excel(xls_file, sheet_name='all_tx', header=None)
listAll = excel_data_df.values.tolist()

#print out the result
#print('All list')
#print(listAll)
                    

#Summary
sum_list = []

for txname, pid, st, et, sec in (listAll) :
    found = False
    for item in sum_list :
        if (item[0] in txname) and (len(str(sec))!=0) and (str(sec) != 'nan'):
            if (item[0] in 'APINQEQG'):
                print('[{}] len={}'.format(sec,len(str(sec))))
            found = True
            item[1] += 1
            item[2] += sec
            item[5] = item[2]/item[1]
            if sec > item[3] :
                item[3] = sec
            if sec < item[4] :
                item[4] = sec 
    if (not found) and (str(sec) != 'nan') :
        new_item = [txname,1,sec,sec,sec,sec]
        sum_list.append(new_item) 

#print out the summary list
print('Summary')
for item in sum_list:
    print(item)


#Insert titles into sheet
title_sum_list = ['TX_NAME','COUNT','SUM','MAX_TIME','MIN_TIME','AVG_TIME']
sum_list.insert(0, title_sum_list)

wb = openpyxl.load_workbook(xls_file)
wb.create_sheet('summary')
ws = wb['summary']
for tx in sum_list:
    ws.append(tx)

wb.save(xls_file) 


