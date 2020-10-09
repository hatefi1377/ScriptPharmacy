from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time , os , shutil

dir_path = os.path.dirname(os.path.realpath(__file__))
url = 'https://darman.tamin.ir/captchaCheck.aspx'


chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : f"{dir_path}\Sites"}
chromeOptions.add_experimental_option("prefs",prefs)
chromedriver = "chromedriver.exe"
driver = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)

driver.get(url)
# login to site
User = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtUID"]').send_keys("2430000000031")
Password = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtPass"]').send_keys("BH47shahrokh")
Chapta = input("Code : ")
SiteChapta = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ASPxCaptcha1_TB_I"]').send_keys(Chapta)
Elem = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnOk"]').click()

# site navigation
Elem = driver.find_element_by_xpath('//*[@id="ctl00_mnuMain_I2i7_T"]/a').click()
month = input("Month : ")
data = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtFromDate"]').send_keys(f"1399/{month}/01")
Elem = driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnSearch"]').click()
driver.implicitly_wait(30)
Main_Window = driver.window_handles[0]
Elem = driver.find_element_by_xpath('//*[@id="btnPrint"]').click()
Elem = WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
Print_window = driver.window_handles[1]
driver.switch_to.window(Print_window)
Elem = driver.find_element_by_xpath('//*[@id="CrystalReportViewer1"]/tbody/tr/td/div/div[1]/table/tbody/tr/td[2]/input').click()
Elem = WebDriverWait(driver, 60 , ).until(EC.number_of_windows_to_be(3))
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

