import os
import glob
from openpyxl import Workbook, load_workbook
from datetime import datetime

files = glob.glob('/mnt/c/Users/narp/AppData/Roaming/bakkesmod/bakkesmod/data/RocketStats/RocketStats_*.txt')


# Collect data

metric_names = ['CurrentDate']

now = datetime.today().strftime('%d-%m-%Y %H:%M')
metric_values = [now]

for f in files:
    metric = f.replace('/mnt/c/Users/narp/AppData/Roaming/bakkesmod/bakkesmod/data/RocketStats/RocketStats_','')
    metric = metric.replace('.txt','')
    metric_names.append(metric)
    with open(f) as ff:
        value = ff.readline()
        try:
            metric_values.append(float(value))
        except ValueError:
            metric_values.append(value)

# Write CSV

first_row = ','.join(metric_names)
second_row = ','.join([str(v) for v in metric_values])
stats_csv = '/home/nat/rl_stats/stats.csv'
if os.path.isfile(stats_csv):
    print('csv file exists, appending...') 
    with open(stats_csv,'a') as csv:
        csv.write('\n')
        csv.write(second_row)        
else:
    print('csv file does not exist, creating...')
    with open(stats_csv,'w') as csv:
        csv.write(first_row)
        csv.write('\n')
        csv.write(second_row)

# Write XLSX

stats_xlsx = '/home/nat/rl_stats/stats.xlsx'
if os.path.isfile(stats_xlsx):
    print('xlsx file exists, adding stats...')
    wb = load_workbook(filename = stats_xlsx)
    ws = wb.active 
    row = 2
    col = 1
    cell_content = ws.cell(row = row,column= col).value
    while cell_content is not None:
        row += 1 
        cell_content = ws.cell(row = row,column= col).value
    for v in metric_values:
        ws.cell(column=col, row=row, value=v)
        col += 1
    wb.save(stats_xlsx)

else:
    print('xlsx file does not exist, creating...')   
    wb = Workbook()
    ws = wb.active  
    row = 1
    col = 1
    for n in metric_names:
        ws.cell(column=col, row=row, value=n)
        col += 1
    row = 2
    col = 1
    for v in metric_values:
        ws.cell(column=col, row=row, value=v)
        col += 1
    wb.save(stats_xlsx)