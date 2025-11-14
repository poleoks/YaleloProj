#%%
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
# import pyperclip

from datetime import datetime, timedelta
# from config import CHROME_PROFILE_PATH
current_time = str(datetime.now().time())
Options=Options()
Options.add_experimental_option("detach", True)

driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)

#%%
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
log_name.send_keys("KEN32027L")

log_name=WebDriverWait(driver, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/aside/div[1]/div[2]/form/div[2]/input"))
    )
# log_name=driver.find_element("id","//input[@id='tb_user_id']")
log_name.send_keys("no48d012")
# log_password=driver.find_element("id","//input[@id='tb_password']")
# log_password.send_keys("no48d012")
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
log_name.send_keys("UGA07527T")

log_name=WebDriverWait(browser, 500).until(
        EC.presence_of_element_located(("xpath","/html/body/div[3]/aside/div[1]/div[2]/form/div[2]/input"))
    )
# log_name=driver.find_element("id","//input[@id='tb_user_id']")
log_name.send_keys("d51063J3")
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
print(f"kenya_df_shape:{ke_fuel.shape}, uganda_df_shape:{ug_fuel.shape}")
print(f"ke_columns:{ke_fuel.columns}, ug_columns:{ug_fuel.columns}")
#%%
# import os
os.remove("C:/Users/Pole Okuttu/Downloads/Transactions.csv")

browser.quit()
# %%
total_fuel=pd.concat([ug_fuel,ke_fuel], ignore_index=True)
total_fuel['Receipt_n_date']= total_fuel['Date'].astype('str') +  total_fuel['Receipt num.'].astype('str')


total_fuel.to_csv('P:/Pertinent Files/Python/scripts/Total Fuel/total_fuel.csv',index=False)
# print(total_fuel.shape)
# print(total_fuel.info())
# print(total_fuel['Currency'].value_counts(normalize=True))
print(total_fuel.tail())
# %%




#send email
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from app_credential import *

# Replace these variables with your actual email details
from_address = personal_email
to_address = 'pokuttu@yalelo.ug'
subject = 'Total Fuel Data'
body = 'Hello, See attached the fuel data captured'

# Create the email
email = MIMEMultipart()
email['From'] = from_address
email['To'] = to_address
email['Subject'] = subject
email.attach(MIMEText(body, 'plain'))

# Convert DataFrame to CSV string
csv_content = total_fuel.to_csv(index=False)

# Add attachment
attachment = MIMEText(csv_content, 'csv')
attachment.add_header('Content-Disposition', 'attachment', filename='total_fuel.csv')
email.attach(attachment)

# Replace the following with Gmail SMTP settings
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_login = personal_email
smtp_password = personal_email_password  # Generate this from your Google Account security settings

# Create the SMTP sender
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(smtp_login, smtp_password)

# Send the email
server.sendmail(from_address, to_address, email.as_string())
#%%
# Close the connection
server.quit()
#%% 
drvr=webdriver.Chrome(service=Service(ChromeDriverManager().install()))

#%%
# groups_path='P:/Pertinent Files/Python/scripts/groups_all.txt'
# drvr.get('https://firstwavegroup.sharepoint.com/sites/FirstWaveDA/')
drvr.get('https://firstwavegroup.sharepoint.com/:f:/r/sites/FirstWaveDA/Shared%20Documents/Yalelo%20Uganda%20(YU)/Recurring%20Reports/YU%20Distribution?csf=1&web=1&e=l2L2k7')

drvr.maximize_window()
time.sleep(2)

#%%
# driver2.find_element("id","//button[@id='tarteaucitronAllDenied2']").click()

WebDriverWait(drvr, 10*60).until(
        EC.presence_of_element_located(("xpath","/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]"))
    ).send_keys("pokuttu@yalelo.ug")

time.sleep(3)
# #%%
WebDriverWait(drvr, 500).until(
        EC.presence_of_element_located(("id","idSIButton9"))
    ).click()
# drvr.find_element("id","idSIButton9").click()

time.sleep(3)
WebDriverWait(drvr, 500).until(
        EC.presence_of_element_located(("id","i0118"))
    ).send_keys("Aligator@1")
# # drvr.find_element("id","i0118").send_keys("Aligator@1")
time.sleep(2)

WebDriverWait(drvr, 500).until(
        EC.presence_of_element_located(("id","idSIButton9"))
    ).click()
# # # drvr.find_element("id","idSIButton9").click()
time.sleep(2)

WebDriverWait(drvr, 500).until(
        EC.presence_of_element_located(("id","idBtn_Back"))
    ).click()
# # drvr.find_element("id","idBtn_Back").click()
time.sleep(2)

# WebDriverWait(drvr, 500).until(
#         EC.presence_of_element_located(("XPATH","//button[@name='Files' and @data-automationid='uploadFileCommand'  and @aria-label='Files' and @aria-disabled='false' and @role='menuitem']"))
#     ).send_keys("C:/Users/Pole Okuttu/Downloads/Total Fuel Payments.xlsx")

# element_xpath = '//*[@data-icon-name="upload"]'
# upload_button = drvr.find_element(By.XPATH, element_xpath)
# upload_button.click()
# time.sleep(5)

# element_xpath = '//span[contains(@class, "ms-ContextualMenu-itemText") and contains(text(), "Files")]'
# files_element = drvr.find_element(By.XPATH, element_xpath)

# # Perform some action, for example, click the element
# files_element.click()

# time.sleep(5)

# excel_file_path = "C:/Users/Pole Okuttu/Downloads/Total Fuel Payments.xlsx"

#%%
print("deleting file from sharepoint now")
try:
    drvr.find_element(By.XPATH,'//div[contains(@id,"checkbox") and contains(@aria-label, "Total Fuel Payments.xlsx")]').click()
    time.sleep(3)

    drvr.find_element(By.XPATH,'//button[@role="menuitem" and @aria-label="More" and span[@data-automationid="splitbuttonprimary"]]').click()

    time.sleep(5)

    drvr.find_element(By.XPATH,'//i[@data-icon-name="delete" and @aria-hidden="true"]').click()
    time.sleep(5)
    # Find the button using the updated XPath
    button_xpath = '//button[@type="button" and @data-automationid="confirmbutton" and .//span[contains(@class, "ms-Button-label") and contains(normalize-space(), "Delete")]]'
    delete_button = drvr.find_element(By.XPATH, button_xpath)

    # # Scroll to the element using JavaScript
    drvr.execute_script("arguments[0].scrollIntoView(true);", delete_button)

    # # Perform actions on the element (e.g., click)
    delete_button.click()
except:
    print("file not found, proceed")
time.sleep(5)
#%%

#Upload file into sharepoint
# Click the button that opens the menu
button_xpath = '//*[@data-icon-name="upload"]'
button_element = drvr.find_element(By.XPATH, button_xpath)
button_element.click()

# Wait for a short time to allow the menu to appear (adjust as needed)
time.sleep(2)

# Select the "Files" option from the menu
files_option_xpath = '//span[contains(@class, "ms-ContextualMenu-itemText") and contains(text(), "Files")]'
files_option_element = drvr.find_element(By.XPATH, files_option_xpath)
files_option_element.click()

# Wait for a short time to allow the file input to become visible (adjust as needed)
time.sleep(2)

# Find the file input element using XPath
file_input_xpath = '//input[@type="file"]'
file_input = drvr.find_element(By.XPATH, file_input_xpath)

# Provide the path to your Excel file
excel_file_path = "C:/Users/Pole Okuttu/Downloads/Total Fuel Payments.xlsx"

# Set the file path in the file input
file_input.send_keys(excel_file_path)

# Wait for a few seconds (just for demonstration purposes)
time.sleep(5)

# Close the browser
drvr.quit()