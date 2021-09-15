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
sheet = wb.active
getInfo(sheet)

for var in range(1,sheet.max_row+1):
	logbook = "date:{date}|ACCMS|project:{project}|".format(date = sheet[var][EndDate].value , project = sheet[var][CMProject].value)
	print(logbook)

