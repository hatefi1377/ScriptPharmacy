# coding=UTF-8
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
import os

dir_path = os.path.dirname(os.path.realpath(__file__))
Dir = dir_path + "\Sites\\mosalah.xls"
workbook = open_workbook(Dir)
worksheet = workbook.sheet_by_index(0)
wb = copy(workbook)
sheet = wb.get_sheet(0)


replacement = {'ميباشد': 'است',
               'نمي باشد': 'نيست'
               }

for i in range(1,worksheet.nrows):
    data = worksheet.cell_value(i, 6)
    if data in replacement.keys():
        sheet.write(i,6,replacement[data])

wb.save(Dir)
