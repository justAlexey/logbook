import time

import openpyxl as opxl
from progress.bar import IncrementalBar

import logbook
import database

print("start")
s = time.monotonic()
lb = logbook.Logbook()
# print("clear 1225")
dif = time.monotonic()
lb.ws1225 = lb.clear1225(lb.ws1225)
<<<<<<< HEAD
#lb.ws1328 = lb.clear1328(lb.ws1328)
=======
lb.ws1328 = lb.clear1328(lb.ws1328)
>>>>>>> origin/main

print("cleared")
progress = IncrementalBar("create logbook...", max = lb.ws1225.max_row-2)

for i in range(2,lb.ws1225.max_row):
    wsrow = lb.ws1225[i]
    if wsrow[lb.tow].value not in database.typeofwork:
        progress.next()
        continue
    data = lb.formateRow(i)
    if data == None:
        progress.next()
        continue
    lb.writeRow(lb.ws1225[i][lb.sbt].value, data)

    progress.next()
lb.save()
progress.finish()
print(f'used time - {time.monotonic()-s}')
