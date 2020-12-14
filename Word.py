from __future__ import print_function
from mailmerge import MailMerge
import xlrd
import Doctors
#from docx2pdf import convert
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
            DOB=value[5],
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
            DOB=value[5],
            code=str(int(value[1]))
        )
        document.write('RESULT/NotFound/ERROR'+str(int(value[1]))+'_'+str(int(value[0]))+'.docx')

#convert('RESULT/', 'RESULT/PDFs/')


# LIBRE_OFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
#
# def convert_to_pdf(input_docx, out_folder):
#     p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
#                out_folder, input_docx])
#     print([LIBRE_OFFICE, '--convert-to', 'pdf', input_docx])
#     p.communicate()
#
#
# sample_doc = 'file.docx'
# out_folder = 'some_folder'
# convert_to_pdf(sample_doc, out_folder)


