import openpyxl as xl

print("hello world")

wb = xl.load_workbook("example.xlsx")

print(type(wb))
print(wb.get_sheet_names())
sheet = wb.get_sheet_by_name("Sheet1")
print(sheet["A1"].value)

i = 1

while sheet.cell(row=i, column=1).value != None:
    print(sheet.cell(row=i, column=1).value, sheet.cell(row=i, column=2).value)
    i += 1


