import os
import openpyxl

currentPath = os.path.dirname(os.path.abspath(__file__))

fileXlsx = currentPath+"/screenshots3.xlsx"

wb = openpyxl.load_workbook(fileXlsx)
ws = wb["Sheet"]

i = 4
for root, dirs, files in os.walk("."):
	dirs.sort()
	for name in dirs:
		# print(name[2:])
		ws["C%d"%(i)] = name
		i += 1

wb.save(fileXlsx)