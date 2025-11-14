
#%%
try:
    file_location = 'C:/Users/Administrator/Documents/Python_Automations/Finance/YU_Fuel_Automated.xlsx'
    
    os.remove(file_location)
    print("file deleted from SharePoint")

except:
    print("file Does not exist")
#%%
#import modules
import selenium
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import pandas as pd 
import sys

from datetime import datetime, timedelta
# from config import CHROME_PROFILE_PATH
current_time = str(datetime.now().time())
sys.path.insert(1,"C:/Users/Administrator/Documents/Python_Automations/")
from credentials import *
#%%
Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)
# Options.add_argument(CHROME_PROFILE_PATH)
chrome_install = ChromeDriverManager().install()

folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

Service = webdriver.ChromeService(chromedriver_path)

driver=webdriver.Chrome(service=Service)

today = datetime.now().strftime('%m/%d/%Y')
#%%
#GET FUEL CONSUMPTION IN KENYA
# groups_path='P:/Pertinent Files/Python/scripts/groups_all.txt'
driver.get('https://www.mytotalfuelcard.com/')
driver.delete_all_cookies()
driver.maximize_window()
time.sleep(5)

#%%
# driver.find_element("id","//button[@id='tarteaucitronAllDenied2']").click()

WebDriverWait(driver, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[5]/div[3]/button[2]"))
    ).click()

log_name=WebDriverWait(driver, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/aside/div[1]/div[2]/form/div[1]/input"))
    )
# log_name=driver.find_element("id","//input[@id='tb_user_id']")
log_name.send_keys(Login_ke)

log_name=WebDriverWait(driver, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/aside/div[1]/div[2]/form/div[2]/input"))
    )
# log_name=driver.find_element("id","//input[@id='tb_user_id']")
log_name.send_keys(Password_ke)
driver.find_element('xpath','/html/body/div[3]/aside/div[1]/div[2]/form/a/div').click()

# %%
transaction_button=WebDriverWait(driver, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/aside/div[1]/nav/ul/li[3]/a"))
    )
transaction_button.click()


customer_filter=WebDriverWait(driver, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/section/div/section/div/table/tbody/tr[1]/td[2]/select"))
    )
customer_filter.click()


select_yu=WebDriverWait(driver, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/section/div/section/div/table/tbody/tr[1]/td[2]/select/option[2]"))
    )
select_yu.click()

bgd = WebDriverWait(driver, 500).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="dp_date_debut"]'))
)

bgd.clear()
time.sleep(1)
bgd.send_keys('01/01/2024')
bgd.send_keys(Keys.ENTER)
time.sleep(1)

edd = WebDriverWait(driver, 500).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="dp_date_fin"]'))
)

edd.clear()
time.sleep(1)
edd.send_keys(today)
edd.send_keys(Keys.ENTER)
time.sleep(1)
#%%
# csv download
download_link=WebDriverWait(driver, 500).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="csvIcon"]'))
    )
driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
download_link.click()

#%%
# Wait for the download to complete (you may need to customize the waiting time)
max_wait_time_seconds = 60*1
current_wait_time = 0

while not os.path.exists("c:/Users/Administrator/Downloads/Transactions.csv") and current_wait_time < max_wait_time_seconds:
    time.sleep(1)
    current_wait_time += 1
# Check if the file exists
ke_fuel=pd.read_csv("c:/Users/Administrator/Downloads/Transactions.csv",sep=";")
print(ke_fuel.shape)
print(ke_fuel.head())
#%%
# import os
os.remove("c:/Users/Administrator/Downloads/Transactions.csv")

driver.quit()
time.sleep(5)
# %%

#GET FUEL CONSUMPTION IN UGANDA
browser=webdriver.Chrome(service=Service)
# groups_path='P:/Pertinent Files/Python/scripts/groups_all.txt'
browser.get('https://www.mytotalfuelcard.com/')
browser.delete_all_cookies()
browser.maximize_window()
time.sleep(5)

#%%
# driver.find_element("id","//button[@id='tarteaucitronAllDenied2']").click()

WebDriverWait(browser, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[5]/div[3]/button[2]"))
    ).click()

log_name=WebDriverWait(browser, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/aside/div[1]/div[2]/form/div[1]/input"))
    )
# log_name=driver.find_element("id","//input[@id='tb_user_id']")
log_name.send_keys(Login_ug)

log_name=WebDriverWait(browser, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/aside/div[1]/div[2]/form/div[2]/input"))
    )
# log_name=driver.find_element("id","//input[@id='tb_user_id']")
log_name.send_keys(Password_ug)
# log_password=driver.find_element("id","//input[@id='tb_password']")
# log_password.send_keys("no48d012")
browser.find_element('xpath','/html/body/div[3]/aside/div[1]/div[2]/form/a/div').click()

# %%
transaction_button=WebDriverWait(browser, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/aside/div[1]/nav/ul/li[3]/a"))
    )
transaction_button.click()


customer_filter=WebDriverWait(browser, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/section/div/section/div/table/tbody/tr[1]/td[2]/select"))
    )
customer_filter.click()


select_yu=WebDriverWait(browser, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/section/div/section/div/table/tbody/tr[1]/td[2]/select/option[2]"))
    )
select_yu.click()


#%%
bgd = WebDriverWait(browser, 500).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="dp_date_debut"]'))
)

# csv download
bgd.clear()
time.sleep(1)
bgd.send_keys('01/01/2024')
bgd.send_keys(Keys.ENTER)
time.sleep(1)

edd = WebDriverWait(browser, 500).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="dp_date_fin"]'))
)

edd.clear()
time.sleep(1)
edd.send_keys(today)
edd.send_keys(Keys.ENTER)
time.sleep(1)
#%%
# csv download
download_link=WebDriverWait(browser, 500).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="csvIcon"]'))
    )
browser.execute_script("arguments[0].scrollIntoView(true);", download_link)
download_link.click()

#%%
# Wait for the download to complete (you may need to customize the waiting time)
max_wait_time_seconds = 60*5
current_wait_time = 0

while not os.path.exists("c:/Users/Administrator/Downloads/Transactions.csv") and current_wait_time < max_wait_time_seconds:
    time.sleep(1)
    current_wait_time += 1
# Check if the file exists
ug_fuel=pd.read_csv("c:/Users/Administrator/Downloads/Transactions.csv",sep=";")
print(ug_fuel.head())
#%%
# import os
os.remove("c:/Users/Administrator/Downloads/Transactions.csv")

browser.quit()
time.sleep(5)