import openpyxl as opxl
from progress.bar import IncrementalBar

import logbook

lb = logbook.Logbook()
# print("clear 1225")
lb.ws1225 = lb.clear(lb.ws1225)
#print("cleared")
#progress = IncrementalBar("create logbook...", max = lb.ws1225.max_row)
for i in range(2,lb.ws1225.max_row):
    data = lb.formateRow(i)
    if data == None:
#        progress.next()
        continue
    lb.writeRow(lb.ws1225[i][lb.sbt].value, data)
#    progress.next
lb.save()
#progress.finish()

