#%%
#import modules

import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
# from selenium.common.exceptions import NoSuchElementException
import os
import pandas as pd 
# from PIL import Image
import pyperclip
#%%

from datetime import date,datetime, timedelta
import calendar
my_date = date.today()


day_of_the_week=calendar.day_name[my_date.weekday()]
day_of_the_week_num=datetime.today().weekday() #0 for Monday, 1 for Tuesday
month_start_date =my_date.replace(day=1).strftime('%m/%d/%Y')
month_end_date = my_date.strftime('%m/%d/%Y')


week_start_date=my_date - timedelta(days=my_date.weekday())
week_end_date=week_start_date+timedelta(days=6)

week_start_date=week_start_date.strftime('%m/%d/%Y')
week_end_date=week_end_date.strftime('%m/%d/%Y')

print(f"Month start date: {month_start_date}")
print(f"Month end date: {month_end_date}")

print(f"week start date: {week_start_date}")
print(f"week end date: {week_end_date}")
#%%
current_time = str(datetime.now().time())

# Options.add_experimental_option("detach", True)

Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)

driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)

#%%
# powerbi login url
driver.get('https://app.powerbi.com/?noSignUpCheck=1')
driver.maximize_window()
# wait for page load
time.sleep(5)
# find email area using xpath

# email = driver.find_element("xpath",'//*[@id="i0116"]')

email=WebDriverWait(driver,5*60).until(
    EC.presence_of_element_located((By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']"))
)
# # send power bi login user
pbi_user='d365@yalelo.ke'
pbi_pass='Kenya@YK'
email.send_keys(pbi_user)
#%% find submit button
submit = WebDriverWait(driver,5*60).until(
    EC.presence_of_element_located((By.ID,'idSIButton9'))
)
# # click for submit button
submit.click()


#%%
# now we need to find password field
password = WebDriverWait(driver,5*60).until(
    EC.presence_of_element_located((By.ID,'i0118'))
)
# then we send our user's password 
password.send_keys(pbi_pass)
# after we find sign in button above
submit = WebDriverWait(driver,5*60).until(
    EC.presence_of_element_located((By.ID,'idSIButton9'))
)
# then we click to submit button
submit.click()
# time.sleep(3)
time.sleep(5)
# to complate the login process we need to click no button from above
# find no button by element
no = WebDriverWait(driver,5*60).until(
    EC.presence_of_element_located((By.ID,'idBtn_Back'))
)
# then we click to no
no.click()
time.sleep(5)
# //*[@id="email"]

#%%
#navigate the report and take screenshot
current_dir= "P:/Pertinent Files/Python/scripts/daily_dispatch_status/"
ss_name ='distribution.png'
ss_path=current_dir+ss_name
# ss_path = os.path.join(current_dir, "/", ss_name)
# print(f"This is the picture path: {ss_path}")

report_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/a9add40f90705eba05c8"

# Navigate to the report URL
driver.get(report_url)
# wait for report load

#%%
#wait for a max of 5 mins until full load
# filter shipping warehouse
# Locate and click the dropdown to make the items visible
dropdown = WebDriverWait(driver, 5*60).until(
        EC.presence_of_element_located((By.XPATH,'//*[@class ="slicer-dropdown-menu" and @aria-label="ShippingWarehouseId"]'))
    )
dropdown.click()
# Add a delay to ensure the dropdown menu is fully loaded
time.sleep(5)  # You may need to adjust the sleep duration based on the page load speed
    # Wait for the search box to be present within the dropdown


ActionChains(dropdown).send_keys(Keys.CONTROL, "f").perform()

# def scroll_within_container(container_xpath, element_xpath):
#     # Locate the container
#     container = driver.find_element(By.XPATH, container_xpath)
    
#     # Scroll until the element is visible
#     while True:
#         try:
#             time.sleep(5)
#             # Try to find the element within the container
#             element = container.find_element(By.XPATH, element_xpath)
#             if element.is_displayed():
#                 return element
#         except:
#             pass  # Element not found, continue scrolling

#         # Scroll the container down
#         driver.execute_script("arguments[0].scrollBy(0, 100);", container)
#         time.sleep(2)  # Wait for the scroll to complete

# # Define the XPath for the container and the element
# container_xpath = "//div[@class='scrollable-container']"  # Replace with your container's XPath
# element_xpath = "//div[@class='slicerItemContainer' and @title='Swift']"

# # Scroll within the container until the element is visible
# element = scroll_within_container(container_xpath, element_xpath)



# # Perform actions with the element if needed
# print("Element found and is visible:", element.text)
time.sleep(10)
# Close the browser
driver.quit()
# #%%
