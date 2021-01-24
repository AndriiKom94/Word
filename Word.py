from __future__ import print_function
from mailmerge import MailMerge
import xlrd
import Doctors
from subprocess import Popen
loc = ('REPORT.xls')

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

for row in range(sheet.nrows):
    value = sheet.row_values(row)
    template = 'UPDATED DX.docx'
    document = MailMerge(template)

    if str(int(value[1])) in Doctors.code:

        document.merge(
            DX=value[2],
            DOS=value[3][1:],
            doctor=Doctors.res[str(int((value[1])))],
            invoice=str(int(value[0])),
            patient=value[4],
            DOB=value[5][1:],
            code=str(int(value[1]))
        )

        document.write('RESULT/'+str(int(value[1]))+'_'+str(int(value[0]))+'.docx')
    else:
        document.merge(
            DX=value[2],
            DOS=value[3][1:],
            doctor='NOT FOUND',
            invoice=str(int(value[0])),
            patient=value[4],
            DOB=value[5][1:],
            code=str(int(value[1]))
        )
        document.write('RESULT/NotFound/ERROR'+str(int(value[1]))+'_'+str(int(value[0]))+'.docx')

