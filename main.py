import openpyxl as opxl

import logbook

lb = logbook.Logbook()
# print("clear 1225")
lb.ws1225 = lb.clear(lb.ws1225)
for i in range(2,lb.ws1225.max_row):
    data = lb.formateRow(i)
    if data == None:
        break
    lb.writeRow(lb.ws1225[i][lb.sbt].value, data)
lb.save()

