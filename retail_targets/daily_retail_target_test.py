# #%%
# #import modules
# import selenium
# import time
# from selenium import webdriver
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.by import By
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.common.exceptions import NoSuchElementException
# import os
# import pandas as pd 
# from PIL import Image
# import pyperclip
# #%%

# from datetime import date,datetime, timedelta
# import calendar
# my_date = date.today()


# day_of_the_week=calendar.day_name[my_date.weekday()]
# day_of_the_week_num=datetime.today().weekday() #0 for Monday, 1 for Tuesday

# print(f"tin: {day_of_the_week_num}")
# #%%
# current_time = str(datetime.now().time())

# # Options.add_experimental_option("detach", True)

# Options=webdriver.ChromeOptions()
# Options.add_experimental_option("detach", True)

# driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)

# #%%
# # powerbi login url
# driver.get('https://office.live.com/start/Excel.aspx')
# driver.maximize_window()
# # wait for page load
# time.sleep(5)
# # find email area using xpath

# # email = driver.find_element("xpath",'//*[@id="i0116"]')

# email=driver.find_element(By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']")
# # # send power bi login user
# pbi_user='d365@yalelo.ke'
# pbi_pass='Kenya@YK'
# email.send_keys(pbi_user)
# #%% find submit button
# submit = driver.find_element("id",'idSIButton9')
# # # click for submit button
# submit.click()

# time.sleep(5)
# #%%
# # now we need to find password field
# password = driver.find_element("id",'i0118')
# # then we send our user's password 
# password.send_keys(pbi_pass)
# # after we find sign in button above
# submit = driver.find_element("id",'idSIButton9')
# # then we click to submit button
# submit.click()
# # time.sleep(10)
# time.sleep(5)
# # to complate the login process we need to click no button from above
# # find no button by element
# no = driver.find_element("id",'idBtn_Back')
# # then we click to no
# no.click()
# time.sleep(5)
# # //*[@id="email"]

# #%%
# #navigate the report and take screenshot
# current_dir= "P:/Pertinent Files/Python/"
# ss_name ='distribution.png'
# ss_path = os.path.join(current_dir, "saved_images/", ss_name)
# print(ss_path)

# # report_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSection2436e75a5249a600b5a0"
# report_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSection57732dac3e37637503de"

# # Navigate to the report URL
# driver.get(report_url)
# # wait for report load

#%%
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

# Replace these with your actual credentials
username = 'pokuttu@yalelo.ug'
password = 'Aligator@1'

# URL of the Excel Online document
excel_url = "https://office.live.com/start/Excel.aspx"

# Start a new instance of the Chrome browser
driver = webdriver.Chrome()

# Open the Excel Online URL
driver.get(excel_url)

# Wait for the page to load
time.sleep(5)

# Find the email input field and enter the username
email_input = driver.find_element("name", "loginfmt")
email_input.send_keys(username)

# Click the "Next" button
driver.find_element("id", "idSIButton9").click()

# Wait for the page to load
time.sleep(5)

# Find the password input field and enter the password
password_input = driver.find_element("name", "passwd")
password_input.send_keys(password)

# Press Enter or click the "Sign in" button
password_input.send_keys(Keys.RETURN)

# Wait for the page to load after login
time.sleep(10)

# Now, you can navigate to the specific Excel sheet and retrieve data using pandas
# For example, let's assume the data is in the first sheet
# driver.get("https://firstwavegroup.sharepoint.com/:x:/r/sites/FirstWaveDA/Shared%20Documents/Date%20-%20Expected%20Date%20of%20Delivery.xlsx?d=wea34f9dec4fd41bf9696f26f9f1e4c62&csf=1&web=1&e=hlODlY&nav=MTVfezAwMDAwMDAwLTAwMDEtMDAwMC0wMDAwLTAwMDAwMDAwMDAwMH0")
driver.get("https://firstwavegroup.sharepoint.com/:x:/r/sites/FirstWaveDA/Shared%20Documents/Date%20-%20Expected%20Date%20of%20Delivery.xlsx")
# driver.get("https://office.live.com/start/Excel.aspx#/worksheet/workbookview")
time.sleep(5)

# Retrieve the table data using pandas
table_data = pd.read_html(driver.page_source)

# Get the DataFrame (assuming the first table on the page contains your data)
df = table_data[0]

# Display the DataFrame
print(df)

# Close the browser window
driver.quit()
