from openpyxl import Workbook
wb = Workbook()
ws = wb.active
l = [1, 2, 3, 4]
row = 1
col = 1
for v in l:
    ws.cell(column=col, row=row, value=v)
    print(v)
    print(row)
    print(col)
    col += 1
wb.save('test.xlsx')
