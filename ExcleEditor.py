# coding=UTF-8
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
import os

dir_path = os.path.dirname(os.path.realpath(__file__))
Dir = dir_path + "\Sites\\khadamat.xls"
WorkbookNew = open_workbook(Dir)
WorksheetNew = WorkbookNew.sheet_by_index(0)
wbNew = copy(WorkbookNew)
sheetNew = wbNew.get_sheet(0)

DirMain = dir_path + "\Sites\\AllPrice.xls"
WorkbookAll = open_workbook(DirMain)
WorksheetAll = WorkbookAll.sheet_by_index(0)
wbAll = copy(WorkbookAll)
sheetAll = wbAll.get_sheet(0)


date = "1399/07/05"

#replacing All price data with khadamat
for i in range(1, WorksheetNew.nrows):
    data = WorksheetNew.cell_value(i, 9)
    if data >= date:
        temp = WorksheetNew.row(i)
        for j in range(1, WorksheetAll.nrows):
            if temp[3].value ==  WorksheetAll.cell_value(j, 0):
                price = (temp[7].value).replace(',',"")
                sheetAll.write(j,1,price)
                sheetAll.write(j,4,temp[4].value)
                FranshizTemp = (temp[8].value)
                Franshiz = ""
                for x in range(5):
                    if (FranshizTemp[x] != '%'):
                        Franshiz += FranshizTemp[x]
                    else:
                        break
                sheetAll.write(j,5,Franshiz)
                sheetAll.write(j,6,temp[11].value)
                sheetAll.write(j,7,temp[5].value)
                sheetAll.write(j,8,temp[6].value)

#replacing All price data with mosalah
Dir = dir_path + "\Sites\\mosalah.xls"
WorkbookNew = open_workbook(Dir)
WorksheetNew = WorkbookNew.sheet_by_index(0)
wbNew = copy(WorkbookNew)
sheetNew = wbNew.get_sheet(0)

Replacement = {
    "نمي باشد": "نيست",
    "ميباشد": "است"
}
for i in range(1, WorksheetNew.nrows):
    temp = WorksheetNew.row(i)
    for j in range(1, WorksheetAll.nrows):
        if temp[0].value == WorksheetAll.cell_value(j, 0):
            sheetAll.write(j,3,str(int(temp[4].value)))
            sheetAll.write(j,11,Replacement[temp[6].value])


#replacing All price data with tamin
Dir = dir_path + "\Sites\\tamin.xls"
WorkbookNew = open_workbook(Dir)
WorksheetNew = WorkbookNew.sheet_by_index(0)
wbNew = copy(WorkbookNew)
sheetNew = wbNew.get_sheet(0)


for i in range(2, WorksheetNew.nrows):
    temp = WorksheetNew.row(i)
    for j in range(1, WorksheetAll.nrows):
        if temp[7].value == WorksheetAll.cell_value(j, 0):
            if (temp[4].value == ""):
                pass
            else:
                price = str(int(temp[4].value))
                sheetAll.write(j,2,price)
            sheetAll.write(j,10,temp[3].value)

wbAll.save(DirMain)





