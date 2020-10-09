from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time , os , shutil

date = input("give me date in format = year/month/days : ")
dir_path = os.path.dirname(os.path.realpath(__file__))
url = 'http://www.esata.ir/web/sakhad/drug'
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : f"{dir_path}\Sites"}
chromeOptions.add_experimental_option("prefs",prefs)
chromedriver = "chromedriver.exe"
driver = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)

driver.get(url)
driver.find_element_by_xpath('//*[@id="_sakhadDrug_INSTANCE_TSgxrZhxtar5_fromDate"]').send_keys(date)
driver.find_element_by_xpath('//*[@id="portlet_sakhadDrug_INSTANCE_TSgxrZhxtar5"]/div/div/div/form/fieldset/div/div[4]/div/input').click()
time.sleep(1)
driver.find_element_by_xpath('//*[@id="portlet_sakhadDrug_INSTANCE_TSgxrZhxtar5"]/div/div/div/form/fieldset/div/div[4]/div/a').click()
time.sleep(1)
filename = max([f"{dir_path}\Sites" + "\\" + f for f in os.listdir(f"{dir_path}\Sites")],key=os.path.getctime)
shutil.move(filename,os.path.join(f"{dir_path}\Sites",r"mosalah.xls"))

url = 'https://mdp.ihio.gov.ir/'
driver.get(url)
driver.find_element_by_xpath('//*[@id="cmbSrchSrvChgStatus-trigger-picker"]').click()
driver.find_element_by_css_selector('#cmbSrchSrvChgStatus-inputEl')
driver.find_element_by_css_selector('#cmbSrchSrvChgStatus-picker-listEl > li:nth-child(6)').click()
driver.find_element_by_xpath('//*[@id="BtnSrchService-btnInnerEl"]').click()
time.sleep(3)
driver.find_element_by_xpath('//*[@id="btnExcelExport2-btnInnerEl"]').click()
time.sleep(3)
filename = max([f"{dir_path}\Sites" + "\\" + f for f in os.listdir(f"{dir_path}\Sites")],key=os.path.getctime)
shutil.move(filename,os.path.join(f"{dir_path}\Sites",r"khadamat.xls"))

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
Elem = driver.find_element_by_xpath('//*[@id="CrystalReportViewer1"]/tbody/tr/td/div/div[1]/table/tbody/tr/td[2]/input').click()
Elem = WebDriverWait(driver, 60 ).until(EC.number_of_windows_to_be(3))
Export_window = driver.window_handles[2]
driver.switch_to.window(Export_window)
Elem = driver.find_element_by_xpath('//*[@id="exportFormatList"]/option[6]').click()
Elem = driver.find_element_by_xpath('//*[@id="submitexport"]').click()
time.sleep(2)
filename = max([f"{dir_path}\Sites" + "\\" + f for f in os.listdir(f"{dir_path}\Sites")],key=os.path.getctime)
shutil.move(filename,os.path.join(f"{dir_path}\Sites",r"tamin.xls"))
driver.close()
driver.switch_to.window(Print_window)
driver.close()
driver.switch_to.window(Main_Window)
Khoroj = driver.find_element_by_xpath('//*[@id="ctl00_wucLogin1_btnLogout"]').click()
driver.close()

print("Downloading is done !!")

