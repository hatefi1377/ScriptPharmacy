from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time , os , shutil


dir_path = os.path.dirname(os.path.realpath(__file__))
url = 'https://mdp.ihio.gov.ir/'
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : f"{dir_path}\Sites"}
chromeOptions.add_experimental_option("prefs",prefs)
chromedriver = "chromedriver.exe"
driver = webdriver.Chrome(executable_path=chromedriver, options=chromeOptions)

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


