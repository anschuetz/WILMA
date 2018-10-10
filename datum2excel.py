#!/usr/bin/python3
from xlrd import open_workbook
from xlutils.copy import copy
from datetime import datetime

def convert_date_to_excel_ordinal(datum) :
    offset = 693594
    current = datum
    n = current.toordinal()
    return (n - offset)


xl_file = r'/home/shares/infodisplay/entschuldigung.xls'
rb = open_workbook(xl_file, formatting_info=True)
wb = copy(rb)
sheet = wb.get_sheet(0)
#sheet.write(2,0,datetime.now().strftime("%d.%m.%Y"))
sheet.write(2,0,convert_date_to_excel_ordinal(datetime.now()))
wb.save(xl_file)

