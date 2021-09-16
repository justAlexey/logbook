import openpyxl as opxl
from rich.table import Table
from rich.console import Console
from datetime import datetime



table = Table(title = "logbook")
table.add_column("data")
table.add_column("User Sign")
table.add_column("CM Project")

def getInfo(data):
	global ToW, UserSign,EndDate,UsedMHR,CMProject
	for var in range(0,data.max_column):
		if data[1][var].value == "ToW":
			ToW=var
		if data[1][var].value == "User Sign":
			UserSign = var
		if data[1][var].value == "End Date":
			EndDate = var
		if data[1][var].value == "Used MHR":
			UsedMHR = var
		if data[1][var].value == "CM Project":
			CMProject = var


def clearEmpty(data, column):
	for var in range(1,data.max_row):
		if data[var][column].value == "":
			data.delete_rows(var,var)
			var = var-1


try:
	apn1225 = opxl.open('test3.xlsx')
except:
	print('no file')
	exit()
logbook = opxl.Workbook()
lbsheet = logbook.active
sheet1225 = apn1225.active
getInfo(sheet1225)



lbsheet.title = sheet1225[2][UserSign].value

#clearEmpty(sheet1225, CMProject)

#print(sheet1225[3][CMProject])

for var in range(3,sheet1225.max_row):
	date = sheet1225[var][EndDate]
#	print(date.value.strftime("%d:%m"))
	table.add_row(sheet1225[var][EndDate].value.strftime(, sheet1225[var][UserSign].value, sheet1225[var][CMProject].value)


console = Console()
console.print(table)
