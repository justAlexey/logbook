import openpyxl as opxl

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



try:
	wb = opxl.open('test3.xlsx')
except:
	print('no file')
	exit()
lb = opxl.Workbook()
lbsheet = lb.active
sheet = wb.active
getInfo(sheet)

lbsheet.title = sheet[2][UserSign].value
#lb.save("logbook.xlsx")
#print("date is {}".format(sheet[3][EndDate].value))
print(sheet[3][EndDate].value.year)