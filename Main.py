# coding=UTF-8
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time, os, shutil

date = input("give me date in format = year/month/days : ")
dir_path = os.path.dirname(os.path.realpath(__file__))
url = 'http://www.esata.ir/web/sakhad/drug'
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": f"{dir_path}\Sites"}
chromeOptions.add_experimental_option("prefs", prefs)
chromedriver = "chromedriver.exe"
driver = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)

driver.get(url)
driver.find_element_by_xpath('//*[@id="_sakhadDrug_INSTANCE_TSgxrZhxtar5_fromDate"]').send_keys(date)
driver.find_element_by_xpath(
    '//*[@id="portlet_sakhadDrug_INSTANCE_TSgxrZhxtar5"]/div/div/div/form/fieldset/div/div[4]/div/input').click()
time.sleep(1)
driver.find_element_by_xpath(
    '//*[@id="portlet_sakhadDrug_INSTANCE_TSgxrZhxtar5"]/div/div/div/form/fieldset/div/div[4]/div/a').click()
time.sleep(1)
filename = max([f"{dir_path}\Sites" + "\\" + f for f in os.listdir(f"{dir_path}\Sites")], key=os.path.getctime)
shutil.move(filename, os.path.join(f"{dir_path}\Sites", r"mosalah.xls"))

url = 'https://mdp.ihio.gov.ir/'
driver.get(url)
driver.find_element_by_xpath('//*[@id="cmbSrchSrvChgStatus-trigger-picker"]').click()
driver.find_element_by_css_selector('#cmbSrchSrvChgStatus-inputEl')
driver.find_element_by_css_selector('#cmbSrchSrvChgStatus-picker-listEl > li:nth-child(6)').click()
driver.find_element_by_xpath('//*[@id="BtnSrchService-btnInnerEl"]').click()
time.sleep(3)
driver.find_element_by_xpath('//*[@id="btnExcelExport2-btnInnerEl"]').click()
time.sleep(3)
filename = max([f"{dir_path}\Sites" + "\\" + f for f in os.listdir(f"{dir_path}\Sites")], key=os.path.getctime)
shutil.move(filename, os.path.join(f"{dir_path}\Sites", r"khadamat.xls"))

url = 'https://darman.tamin.ir/captchaCheck.aspx'
driver.get(url)
# login to site
User = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtUID"]').send_keys("2430000000031")
Password = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtPass"]').send_keys("BH47shahrokh")
Chapta = input("Code : ")
SiteChapta = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ASPxCaptcha1_TB_I"]').send_keys(Chapta)
Elem = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnOk"]').click()

# site navigation
Elem = driver.find_element_by_xpath('//*[@id="ctl00_mnuMain_I2i7_T"]/a').click()
data = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtFromDate"]').send_keys(date)
Elem = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnSearch"]').click()
driver.implicitly_wait(30)
Main_Window = driver.window_handles[0]
Elem = driver.find_element_by_xpath('//*[@id="btnPrint"]').click()
Elem = WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
Print_window = driver.window_handles[1]
driver.switch_to.window(Print_window)
Elem = driver.find_element_by_xpath(
    '//*[@id="CrystalReportViewer1"]/tbody/tr/td/div/div[1]/table/tbody/tr/td[2]/input').click()
Elem = WebDriverWait(driver, 60).until(EC.number_of_windows_to_be(3))
Export_window = driver.window_handles[2]
driver.switch_to.window(Export_window)
Elem = driver.find_element_by_xpath('//*[@id="exportFormatList"]/option[6]').click()
Elem = driver.find_element_by_xpath('//*[@id="submitexport"]').click()
time.sleep(2)
filename = max([f"{dir_path}\Sites" + "\\" + f for f in os.listdir(f"{dir_path}\Sites")], key=os.path.getctime)
shutil.move(filename, os.path.join(f"{dir_path}\Sites", r"tamin.xls"))
driver.close()
driver.switch_to.window(Print_window)
driver.close()
driver.switch_to.window(Main_Window)
Khoroj = driver.find_element_by_xpath('//*[@id="ctl00_wucLogin1_btnLogout"]').click()
driver.close()

print("Downloading is done !!")

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

# replacing All price data with khadamat
for i in range(1, WorksheetNew.nrows):
    data = WorksheetNew.cell_value(i, 9)
    if data >= date:
        temp = WorksheetNew.row(i)
        for j in range(1, WorksheetAll.nrows):
            if temp[3].value == WorksheetAll.cell_value(j, 0):
                price = temp[7].value.replace(',', "")
                sheetAll.write(j, 1, price)
                sheetAll.write(j, 4, temp[4].value)
                FranshizTemp = temp[8].value
                Franshiz = ""
                for x in range(5):
                    if FranshizTemp[x] != '%':
                        Franshiz += FranshizTemp[x]
                    else:
                        break
                sheetAll.write(j, 5, Franshiz)
                sheetAll.write(j, 6, temp[11].value)
                sheetAll.write(j, 7, temp[5].value)
                sheetAll.write(j, 8, temp[6].value)

# replacing All price data with mosalah
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
            sheetAll.write(j, 3, str(int(temp[4].value)))
            sheetAll.write(j, 11, Replacement[temp[6].value])

# replacing All price data with tamin
Dir = dir_path + "\Sites\\tamin.xls"
WorkbookNew = open_workbook(Dir)
WorksheetNew = WorkbookNew.sheet_by_index(0)
wbNew = copy(WorkbookNew)
sheetNew = wbNew.get_sheet(0)

for i in range(2, WorksheetNew.nrows):
    temp = WorksheetNew.row(i)
    for j in range(1, WorksheetAll.nrows):
        if temp[7].value == WorksheetAll.cell_value(j, 0):
            if temp[4].value == "":
                pass
            else:
                price = str(int(temp[4].value))
                sheetAll.write(j, 2, price)
            sheetAll.write(j, 10, temp[3].value)

wbAll.save(DirMain)

print("Your File is ready ! ( Press Enter To Close This Window )")
input()
