import xlrd
loc = ('DOCTORS.xls')

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
doctor = []
code = []
for i in range(sheet.nrows):
    doctor.append(sheet.cell_value(i, 0))
    code.append(str(int(sheet.cell_value(i, 1))))


res = {}
for key in code:
    for value in doctor:
        res[key] = value
        doctor.remove(value)
        break

print(res)