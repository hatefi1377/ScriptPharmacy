# coding=UTF-8
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
import os

dateIn = input("give me date in format = year/month/days : ")

date = dateIn[:4] + '/' + dateIn[5:7] + '/' + dateIn[8:]
print(date)