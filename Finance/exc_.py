
from credentials import *
#%%
#%%
# from urllib3.contrib import pyopenssl
from excel_cleaning import receivables
import glob
import os
file_location='c:/Users/Administrator/Downloads/Customer aging report.xlsx'
try:
    for file in glob.glob(os.path.join("c:/Users/Administrator/Downloads","*.xlsx*")):
        os.remove(file)
    for xl in glob.glob(os.path.join("C:/Users/Administrator/Documents/Python_Automations/Finance/","*.xlsx*")):
        os.remove(xl)
                          
    
    # os.remove(file_location)
    print("File deleted")
except:
    print("File Does not Exist")

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
import os
import pandas as pd 
from PIL import Image
import pyperclip
import datetime as dt
from datetime import datetime#, timedelta
import sys
# sys.path.insert(0,'P:/Pertinent Files/Python/scripts/Finance/')

date_today=datetime.today().day
month_num=datetime.today().month
year=datetime.today().year
# Format the date as "mm/dd/yyyy" using strftime()
formatted_date = datetime.today().strftime('%m/%d/%Y')
date_ar= datetime.today().strftime('%Y-%m-%d')
print(formatted_date)

#%%
# from config import CHROME_PROFILE_PATH
current_time = str(datetime.now().time())
Options=Options()
Options.add_experimental_option("detach", True)

#%%
driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
driver.get('https://app.powerbi.com/?noSignUpCheck=1')

driver.maximize_window()
# wait for page load
time.sleep(5)
# find email area using xpath

# email = driver.find_element("xpath",'//*[@id="i0116"]')

email=driver.find_element(By.XPATH,'//*[@aria-label="Enter your email or phone"]')
# # send power bi login user
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

driver.get('https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSection2539a1060abd371bc1e9?experience=power-bi&clientSideAuth=0')

# time.sleep(10)
# export_path=driver.find_element(By.XPATH,"//div[contains(@class, 'pivotTableCellWrap') and contains(text(), 'CustomerPriorityClassificationGroupCode')]").click()
# time.sleep(2)


export_path=WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'pivotTableCellWrap') and contains(text(), 'CustomerPriorityClassificationGroupCode')]"))
)

export_path.click()
time.sleep(2)
driver.find_element(By.XPATH, "//*[@data-testid='visual-more-options-btn']").click()
time.sleep(2)
driver.find_element(By.XPATH, '//*[@data-testid="pbimenu-item.Export data"]').click()
time.sleep(2)

driver.find_element(By.XPATH, '//*[@data-testid="export-btn" ]').click()
time.sleep(5)
print("Done!!")


driver.quit()

time.sleep(5)
# %%

browser=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)


#GET AR REPORT FOR KENYA

browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yk&mi=Output%3ACustAgingBalance')
browser.maximize_window()
# time.sleep(5)
time.sleep(5)
# find email area using xpath

# email = browser.find_element("xpath",'//*[@id="i0116"]')

email=browser.find_element(By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']")
# # send power bi login user
pbi_user='d365@yalelo.ke'
pbi_pass='Kenya@YK'
email.send_keys(pbi_user)
#%% find submit button
submit = browser.find_element("id",'idSIButton9')
# # click for submit button
submit.click()

time.sleep(5)
#%%
# now we need to find password field
password = browser.find_element("id",'i0118')
# then we send our user's password 
password.send_keys(pbi_pass)
# after we find sign in button above
submit = browser.find_element("id",'idSIButton9')
# then we click to submit button
submit.click()
# time.sleep(10)
time.sleep(5)
# to complate the login process we need to click no button from above
# find no button by element
no = browser.find_element("id",'idBtn_Back')
# then we click to no
no.click()
time.sleep(5)
# //*[@id="email"]
#%%

aging_date=WebDriverWait(browser, 500).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="SysOperationTemplateForm_2_Fld3_1_input"]'))
)
aging_date.clear()
aging_date.send_keys(formatted_date)
time.sleep(2)

balance_date=browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld4_1_input"]')
balance_date.clear()
balance_date.send_keys(formatted_date)
time.sleep(2)

# criteria filter selection
browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_input"]').click()
time.sleep(2)

browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_list_item0"]').click()
time.sleep(2)


#aging period definition
browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld6_1"]/div/div').click()
time.sleep(2)

browser.find_element(By.XPATH,'//input[@value="YKCustomerAging"]').click()
time.sleep(2)

#currency filter
browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld7_1"]/div/div[2]/div').click()
time.sleep(2)

browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld7_1_list_item0"]').click()
time.sleep(2)

#negative balance
try:
    browser.find_element(By.XPATH, "//*[@id='SysOperationTemplateForm_2_Fld15_1']//span[@class='toggle-value' and @title='No']").click()
    toggle_element = browser.find_element(By.XPATH, "//span[@id='SysOperationTemplateForm_2_Fld15_1_toggle']")
    toggle_element.click()  # Simulate a click
except:
    print("switched on")
    
# browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld15_1_toggle"]').click()
# time.sleep(2)

#ok button to load report
browser.find_element(By.XPATH, '//*[@id="SysOperationTemplateForm_2_CommandButton"]').click()


#wait for load and press export button

search_xpath = '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportMenuButton_label"]'
search_button = WebDriverWait(browser, 500).until(
    EC.presence_of_element_located((By.XPATH, search_xpath))
)

search_button.click()
time.sleep(2)

browser.find_element(By.XPATH, '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportToExcelButton_label"]').click()
time.sleep(10)

country_ar_extract=file_location
customer_details_extract="c:/Users/Administrator/Downloads/data.xlsx"


kenya_ar, customer_ky=receivables(country_ar_extract,customer_details_extract)

kenya_ar.to_excel('C:/Users/Administrator/Documents/Python_Automations/Finance/ky_Report.xlsx')
customer_ky.to_excel('C:/Users/Administrator/Documents/Python_Automations/Finance/customer_ky.xlsx')


print(kenya_ar.head())
print(kenya_ar.columns)


os.remove(file_location)
browser.quit()

time.sleep(5)

#%%

#***UGANDA****#
drvr=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
drvr.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=Output%3ACustAgingBalance')
drvr.maximize_window()
# time.sleep(5)
time.sleep(5)
# find email area using xpath

# email = drvr.find_element("xpath",'//*[@id="i0116"]')

email=drvr.find_element(By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']")
# # send power bi login user
pbi_user='d365@yalelo.ke'
pbi_pass='Kenya@YK'
email.send_keys(pbi_user)
#%% find submit button
submit = drvr.find_element("id",'idSIButton9')
# # click for submit button
submit.click()

time.sleep(5)
#%%
# now we need to find password field
password = drvr.find_element("id",'i0118')
# then we send our user's password 
password.send_keys(pbi_pass)
# after we find sign in button above
submit = drvr.find_element("id",'idSIButton9')
# then we click to submit button
submit.click()
# time.sleep(10)
time.sleep(5)
# to complate the login process we need to click no button from above
# find no button by element
no = drvr.find_element("id",'idBtn_Back')
# then we click to no
no.click()
time.sleep(5)
# //*[@id="email"]
#%%

aging_date=WebDriverWait(drvr, 500).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="SysOperationTemplateForm_2_Fld3_1_input"]'))
)
aging_date.clear()
aging_date.send_keys(formatted_date)
time.sleep(2)

balance_date=drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld4_1_input"]')
balance_date.clear()
balance_date.send_keys(formatted_date)
time.sleep(2)

# criteria filter selection
drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_input"]').click()
time.sleep(2)

drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_list_item0"]').click()
time.sleep(2)




#currency filter
currency_filter=drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld7_1_input"]')
currency_filter.clear()

time.sleep(2)
currency_filter.send_keys("Accounting currency")

time.sleep(2)

# drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld7_1_list_item0"]').click()
# time.sleep(2)

#negative balance
try:
    drvr.find_element(By.XPATH, "//*[@id='SysOperationTemplateForm_2_Fld15_1']//span[@class='toggle-value' and @title='No']").click()
    toggle_element = drvr.find_element(By.XPATH, "//span[@id='SysOperationTemplateForm_2_Fld15_1_toggle']")
    toggle_element.click()  # Simulate a click
except:
    print("switched on")
    

# drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld15_1_toggle"]').click()
# time.sleep(2)
#aging period definition
drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld6_1"]/div/div').click()
time.sleep(2)

drvr.find_element(By.XPATH,'//input[@value="AgingYU"]').click()
time.sleep(2)
# ok button to load report
drvr.find_element(By.XPATH, '//*[@id="SysOperationTemplateForm_2_CommandButton"]').click()


#wait for load and press export button

search_xpath = '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportMenuButton_label"]'
search_button = WebDriverWait(drvr, 500).until(
    EC.presence_of_element_located((By.XPATH, search_xpath))
)

search_button.click()
time.sleep(2)

drvr.find_element(By.XPATH, '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportToExcelButton_label"]').click()
time.sleep(10)

country_ar_extract=file_location


customer_details_extract="c:/Users/Administrator/Downloads/data.xlsx"


uganda_ar, customer_ug=receivables(country_ar_extract,customer_details_extract)

uganda_ar.to_excel('C:/Users/Administrator/Documents/Python_Automations/Finance/ug_Report.xlsx')
customer_ug.to_excel('C:/Users/Administrator/Documents/Python_Automations/Finance/customer_ug.xlsx')


print(uganda_ar.head())
print(uganda_ar.columns)

drvr.quit()
time.sleep(5)

#%%
import pandas as pd
report_path="C:/Users/Administrator/Documents/Python_Automations/Finance/AR_Report.xlsx"


uganda_ar=pd.read_excel("C:/Users/Administrator/Documents/Python_Automations/Finance/ug_Report.xlsx")
kenya_ar=pd.read_excel("C:/Users/Administrator/Documents/Python_Automations/Finance/ky_Report.xlsx")
customer_ug=pd.read_excel('C:/Users/Administrator/Documents/Python_Automations/Finance/customer_ug.xlsx')
customer_ky=pd.read_excel('C:/Users/Administrator/Documents/Python_Automations/Finance/customer_ky.xlsx')



AR_Standing_status=pd.concat([uganda_ar,kenya_ar])
AR_Standing_status.to_excel(report_path, sheet_name="AR_Sheet", index=False)

print(AR_Standing_status.head())

#%%

AR_Standing_status = pd.ExcelWriter('C:/Users/Administrator/Documents/Python_Automations/Finance/AR_Report.xlsx')

# Write kenya_ar data to a sheet named "Kenya AR"
pd.concat([uganda_ar.iloc[:,:2],kenya_ar.iloc[:,:2]]).to_excel(AR_Standing_status, sheet_name='Consolidated Summary By Sales Channel',index=False)

uganda_ar.to_excel(AR_Standing_status, sheet_name='By Sales Channel - Ug',index=False)
kenya_ar.to_excel(AR_Standing_status, sheet_name='By Sales Channel - Ky',index=False)

# Write customer_ky data to a separate sheet named "Customer KY"
pd.concat([customer_ug.iloc[:,:],customer_ky.iloc[:,:]])[['CustomerAccount','Description','CustomerPriorityClassificationGroupCode',
                                                          'Name','Balance_as_at:{date_ar}']].to_excel(AR_Standing_status, sheet_name='Customer Details',index=False)
