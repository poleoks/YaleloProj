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
# from selenium.common.exceptions import NoSuchElementException
# from config import CHROME_PROFILE_PATH
import os
import pandas as pd 
#%%

from datetime import date,datetime, timedelta
import calendar
my_date = date.today() #-timedelta(days=1)


day_of_the_week=calendar.day_name[my_date.weekday()]
day_of_the_week_num=datetime.today().weekday() #0 for Monday, 1 for Tuesday
month_start_date =my_date.replace(day=1).strftime('%m/%d/%Y')
month_end_date = my_date.strftime('%m/%d/%Y')



this_week_start=my_date - timedelta(days=my_date.weekday())
this_week_end=this_week_start+timedelta(days=6)
sd_date=this_week_start#-timedelta(days=7)
ed_date=this_week_end#- timedelta(days=7)

this_week_start=sd_date.strftime('%m/%d/%Y')
this_week_end=ed_date.strftime('%m/%d/%Y')

# this_week_start=this_week_start.strftime('%m/%d/%Y')
# this_week_end=this_week_end.strftime('%m/%d/%Y')


print(f"My date: {my_date.strftime('%m/%d/%Y')},{my_date.weekday()}")
print(f"Month start date: {month_start_date}")
print(f"Month end date: {month_end_date}")

print(f"week start date: {this_week_start}")
print(f"week end date: {this_week_end}")
#%%
current_time = str(datetime.now().time())
from whatsapp_configuration import CHROME_PROFILE_PATH

#INSTANTIATE WHATSAPP
Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)

chrome_install = ChromeDriverManager().install()

folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

Service = webdriver.ChromeService(chromedriver_path)
driver=webdriver.Chrome(service=Service)

driver.delete_all_cookies()
#%%
# powerbi login url
driver.get('https://app.powerbi.com/?noSignUpCheck=1')
driver.maximize_window()
# wait for page load
time.sleep(5)
# find email area using xpath

# email = driver.find_element("xpath",'//*[@id="i0116"]')

email=driver.find_element(By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']")
# # send power bi login user
import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/')
from credentials import  *
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
# time.sleep(3)
time.sleep(5)
# to complate the login process we need to click no button from above
# find no button by element
no = driver.find_element("id",'idBtn_Back')
# then we click to no
no.click()
time.sleep(5)
# //*[@id="email"]

#%%
# #navigate the report and take screenshot
current_dir= "C:/Users/Administrator/Documents/Python_Automations/Distribution"

ss_name ='distribution.png'
ss_path=current_dir+ss_name
# ss_path = os.path.join(current_dir, "/", ss_name)
print(f"This is the picture path: {ss_path}")

# report_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSection2436e75a5249a600b5a0"
report_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSectionb07695af0594eb707bed"


# 
driver.get(report_url)
# wait for report load

#%%
#wait for a max of 5 mins until full load
#filter date
if day_of_the_week_num==0:
    # select the last week calendar performance from the date filter
    ed=WebDriverWait(driver, 5*60).until(
            EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"End date") and @aria-description="Enter date in M/d/yyyy format"]'))
        )
    ed.clear()
    ed.send_keys(month_end_date)
    ed.send_keys(Keys.ENTER)

    driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()

    time.sleep(3)
    
    sd=WebDriverWait(driver, 5*60).until(
        EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"Start date") and @aria-description="Enter date in M/d/yyyy format"]'))
    )
    sd.clear()
    sd.send_keys(month_start_date)
    sd.send_keys(Keys.ENTER)

    time.sleep(3)
    
    
    driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()
    #Expand view to full screen
    driver.find_element(By.XPATH,"//button[contains(@class, 'mat-menu-trigger') and contains(@aria-label, 'View')]").click()
    expand_button = driver.find_element(By.XPATH, "//button[contains(@class, 'appBarMatMenu') and contains(@title, 'Open in full-screen mode')]")

    # Click the expand button
    expand_button.click()

    time.sleep(5)

    try:
        driver.find_element(By.XPATH,'//*[@aria-label="Show/hide filter pane" and @aria-expanded="true"]').click()
        print("button pressed to hide filter pane")
    except:
        print('Filter already hidden')
    time.sleep(3)
    # Take a screenshot of the report and save it to a file
    # time.sleep(5)
    driver.save_screenshot(ss_path)
    print("successfully taken screenshot!!!")

    time.sleep(3)
    driver.quit()

else:# day_of_the_week_num == 6:
    ed=WebDriverWait(driver, 5*60).until(
            EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"End date") and @aria-description="Enter date in M/d/yyyy format"]'))
        )
    ed.clear()
    ed.send_keys(this_week_end)
    ed.send_keys(Keys.ENTER)

    driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()
    # select the last week calendar performance from the date filter
    time.sleep(3)
    
    
    sd=WebDriverWait(driver, 5*60).until(
        EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"Start date") and @aria-description="Enter date in M/d/yyyy format"]'))
    )
    sd.clear()
    sd.send_keys(this_week_start)
    sd.send_keys(Keys.ENTER)
    
    driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()
    # select the last week calendar performance from the date filter
    time.sleep(3)
    #Expand view to full screen
    driver.find_element(By.XPATH,"//button[contains(@class, 'mat-menu-trigger') and contains(@aria-label, 'View')]").click()
    expand_button = driver.find_element(By.XPATH, "//button[contains(@class, 'appBarMatMenu') and contains(@title, 'Open in full-screen mode')]")

    # Click the expand button
    expand_button.click()

    try:
        driver.find_element(By.XPATH,'//*[@aria-label="Show/hide filter pane" and @aria-expanded="true"]').click()
        print("button pressed to hide filter pane")
    except:
        print('Filter already hidden')
        
    time.sleep(5)
    # Take a screenshot of the report and save it to a file
    # time.sleep(5)
    driver.save_screenshot(ss_path)
    print("successfully taken screenshot!!!")

    time.sleep(5)
    driver.quit()