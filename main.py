from openpyxl import load_workbook

try:
	wb = load_workbook('./test.xlsx')
except:
	print('no file')
	exit()
print(wb.get_sheet_names)
