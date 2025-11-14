#%%
#import modules
import time
import warnings
warnings.filterwarnings("ignore")
import selenium
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
#%%

from datetime import date,datetime, timedelta
import calendar
my_date = date.today()

day_of_the_week=calendar.day_name[my_date.weekday()]
day_of_the_week_num=datetime.today().weekday() #0 for Monday, 1 for Tuesday
#%%

current_time = str(datetime.now().time())


Options=Options()
Options.add_experimental_option("detach", True)

driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)

#%%
# powerbi login url
driver.get('https://app.powerbi.com/?noSignUpCheck=1')
driver.maximize_window()
# wait for page load
time.sleep(5)
# find email area using xpath
email = driver.find_element("xpath",'//*[@id="i0116"]')
# # send power bi login user
pbi_user='d365@yalelo.ke'
pbi_pass='Kenya@YK'
email.send_keys(pbi_user)
#%% find submit button
submit = driver.find_element("id",'idSIButton9')
# # click for submit button
submit.click()
time.sleep(5)

#%%
# now we need to find password field
password = driver.find_element("id",'i0118')
# then we send our user's password 
password.send_keys(pbi_pass)
# after we find sign in button above
submit = driver.find_element("id",'idSIButton9')
# then we click to submit button
submit.click()
# time.sleep(10)
time.sleep(5)
# to complate the login process we need to click no button from above
# find no button by element
no = driver.find_element("id",'idBtn_Back')
# then we click to no
no.click()
time.sleep(5)
# %%
#navigate the report and take screenshot
current_dir= "C:/Users/Administrator/Documents/Python_Automations/Distribution/"
ss_name ='distribution.png'
ss_path = os.path.join(current_dir, "saved_images/", ss_name)
print(ss_path)

# report_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSection2436e75a5249a600b5a0"
report_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/327e7fb0-8c3c-4596-bdd3-98e41603c2a6/ReportSection2e1d7001a05a5136ed5d?experience=power-bi"

# Navigate to the report URL
driver.get(report_url)
# wait for report load
time.sleep(5)
# Take a screenshot of the report and save it to a file
driver.save_screenshot(ss_path)
print("successfully taken screenshot!!!")

time.sleep(5)
# driver.quit()