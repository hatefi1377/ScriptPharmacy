from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time , os , shutil

date = "1399/05/01"
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
