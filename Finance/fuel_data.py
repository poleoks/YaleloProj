from credentials import *
#%%
try:
    file_location = 'P:/Pertinent Files/Python/scripts/Total Fuel/YU_Fuel_Automated.xlsx'
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
from PIL import Image
import pyperclip

from datetime import datetime, timedelta
# from config import CHROME_PROFILE_PATH
current_time = str(datetime.now().time())
Options = Options()
Options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)

#%%
#GET FUEL CONSUMPTION IN KENYA
# groups_path='P:/Pertinent Files/Python/scripts/groups_all.txt'
driver.get('https://www.mytotalfuelcard.com/')
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


#%%
# csv download
download_link=WebDriverWait(driver, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/section/div/section/div/div[3]/div[2]"))
    )
download_link.click()

#%%
# Wait for the download to complete (you may need to customize the waiting time)
max_wait_time_seconds = 60*1
current_wait_time = 0

while not os.path.exists("C:/Users/Pole Okuttu/Downloads/Transactions.csv") and current_wait_time < max_wait_time_seconds:
    time.sleep(1)
    current_wait_time += 1
# Check if the file exists
ke_fuel=pd.read_csv("C:/Users/Pole Okuttu/Downloads/Transactions.csv",sep=";")
print(ke_fuel.shape)
print(ke_fuel.head())
#%%
# import os
os.remove("C:/Users/Pole Okuttu/Downloads/Transactions.csv")

driver.quit()
time.sleep(5)
# %%

#GET FUEL CONSUMPTION IN UGANDA
browser=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
# groups_path='P:/Pertinent Files/Python/scripts/groups_all.txt'
browser.get('https://www.mytotalfuelcard.com/')
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
# csv download
download_link=WebDriverWait(browser, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/section/div/section/div/div[3]/div[2]"))
    )
download_link.click()

#%%
# Wait for the download to complete (you may need to customize the waiting time)
max_wait_time_seconds = 60*1
current_wait_time = 0

while not os.path.exists("C:/Users/Pole Okuttu/Downloads/Transactions.csv") and current_wait_time < max_wait_time_seconds:
    time.sleep(1)
    current_wait_time += 1
# Check if the file exists
ug_fuel=pd.read_csv("C:/Users/Pole Okuttu/Downloads/Transactions.csv",sep=";")
#%%
# import os
os.remove("C:/Users/Pole Okuttu/Downloads/Transactions.csv")

browser.quit()
time.sleep(5)