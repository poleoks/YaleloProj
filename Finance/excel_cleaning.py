<<<<<<< HEAD
import pandas as pd
import datetime as dt
date_today=dt.date.today()
print()
country_ar_extract='C:/Users/Administrator/Downloads/Customer aging report.xlsx'
customer_details_extract='C:/Users/Administrator/Downloads/data.xlsx'

bb=f"As at: {date_today} [ugx]"


df=pd.read_excel(country_ar_extract,skiprows=11,skipfooter=1)
df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
df.rename(columns={ df.columns[3]:bb,df.columns[4]:"90+ days [ugx]",df.columns[5]:"90 days [ugx]",df.columns[6]:"60 days [ugx]",
                   df.columns[7]:"30 days [ugx]",df.columns[8]:"Current Date [ugx]"}, inplace = True)
customer_details=df
#%%
ar=pd.read_excel(customer_details_extract)
print(ar.head(2))


# full_ar=ar[['CustomerAccount', 'Description', 'CustomerPriorityClassificationGroupCode']].join(customer_details, how='left')

full_ar=pd.merge(ar,customer_details, right_on='Account', left_on='CustomerAccount', how='left')
print(full_ar.head())


sales_channel=full_ar.drop(columns=['Account','CustomerAccount','Name','Customer group'])

print(sales_channel.columns)


country_ar_summary=sales_channel.groupby(['Description', 'CustomerPriorityClassificationGroupCode']).sum()
country_ar_summary=country_ar_summary[country_ar_summary.iloc[:,3]>0]

full_ar=full_ar.groupby(['CustomerAccount', 'Description', 'CustomerPriorityClassificationGroupCode',
                                'Account','CustomerAccount','Name','Customer group']).sum()
print(country_ar_summary.head(),full_ar.head())

# return country_ar_summary, full_ar
=======

from credentials import *
#%%
#%%
# from urllib3.contrib import pyopenssl
# from excel_cleaning_rework import receivables
from xlsx_cleaning import ug_extract,ke_extract

#%%
import glob
import os

#%%
file_location='C:/Users/Pole Okuttu/Downloads/Customer aging report.xlsx'
# try:
#     for file in glob.glob(os.path.join("C:/Users/Pole Okuttu/Downloads","*.xlsx*")):
#         os.remove(file)
#     for xl in glob.glob(os.path.join("P:/Pertinent Files/Python/scripts/Finance/","*.xlsx*")):
#         os.remove(xl)
#     # os.remove(file_location)
#     print("File deleted")
# except:
#     print("File Does not Exist")

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
# from PIL import Image
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
# driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
# driver.get('https://app.powerbi.com/?noSignUpCheck=1')

# driver.maximize_window()
# # wait for page load
# time.sleep(5)
# # find email area using xpath

# # email = driver.find_element("xpath",'//*[@id="i0116"]')
# email=WebDriverWait(driver, 500).until(
#     EC.presence_of_element_located((By.XPATH,'//*[@aria-label="Enter your email or phone"]'))
# )
# # email=driver.find_element(By.XPATH,'//*[@aria-label="Enter your email or phone"]')
# # # send power bi login user
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
# time.sleep(5)
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

# driver.get('https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSection2539a1060abd371bc1e9?experience=power-bi&clientSideAuth=0')

# # time.sleep(10)
# # export_path=driver.find_element(By.XPATH,"//div[contains(@class, 'pivotTableCellWrap') and contains(text(), 'CustomerPriorityClassificationGroupCode')]").click()
# # time.sleep(2)


# export_path=WebDriverWait(driver, 500).until(
#     EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'pivotTableCellWrap') and contains(text(), 'CustomerPriorityClassificationGroupCode')]"))
# )

# export_path.click()
# time.sleep(2)
# driver.find_element(By.XPATH, "//*[@data-testid='visual-more-options-btn']").click()
# time.sleep(2)
# driver.find_element(By.XPATH, '//*[@data-testid="pbimenu-item.Export data"]').click()
# time.sleep(2)

# driver.find_element(By.XPATH, '//*[@data-testid="export-btn" ]').click()
# time.sleep(5)
# driver.quit()
# print("Customer Details Extracted From PowerBI")
# %%

browser=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)


#GET AR REPORT FOR KENYA

browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yk&mi=Output%3ACustAgingBalance')
browser.maximize_window()
# time.sleep(5)
time.sleep(5)
# find email area using xpath

# email = browser.find_element("xpath",'//*[@id="i0116"]')
# email=browser.webdr

email=WebDriverWait(browser, 500).until(
    EC.presence_of_element_located((By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']"))
)
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
# country_ar_extract='C:/Users/Pole Okuttu/Downloads/Customer aging report.xlsx'
customer_details_extract='C:/Users/Pole Okuttu/Downloads/data.xlsx'

kenya_ar, customer_ky=ke_extract(country_ar_extract,customer_details_extract)


kenya_ar.to_excel('P:/Pertinent Files/Python/scripts/Finance/ky_Report.xlsx')
customer_ky.to_excel('P:/Pertinent Files/Python/scripts/Finance/customer_ky.xlsx')



print(kenya_ar.head())
print(kenya_ar.columns)


browser.quit()

time.sleep(5)

#%%

#***UGANDA****#
#%%
import os
try:
    for x in glob.glob(os.path.join("C:/Users/Pole Okuttu/Downloads","Customer aging report*")):
        os.remove(x)
except:
    print('does not exist')

#%%
drvr=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
drvr.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=Output%3ACustAgingBalance')
drvr.maximize_window()



time.sleep(5)
# find email area using xpath

# email = drvr.find_element("xpath",'//*[@id="i0116"]')

#%%
# email = drvr.find_element("xpath",'//*[@id="i0116"]')

# email=drvr.find_element(By.XPATH,'//*[@aria-label="Enter your email or phone"]')

email=WebDriverWait(drvr, 500).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]'))
)

#%%
# email=drvr.find_element(By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']")
# # # send power bi login user
# pbi_user='d365@yalelo.ke'
# pbi_pass='Kenya@YK'
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

drvr.quit()
time.sleep(5)
country_ar_extract=file_location


uganda_ar, customer_ug=ug_extract(country_ar_extract,customer_details_extract)

uganda_ar.to_excel('P:/Pertinent Files/Python/scripts/Finance/ug_Report.xlsx')
customer_ug.to_excel('P:/Pertinent Files/Python/scripts/Finance/customer_ug.xlsx')


print(uganda_ar.head())
print(uganda_ar.columns)


#%%
import pandas as pd


# uganda_ar=pd.read_excel("P:/Pertinent Files/Python/scripts/Finance/ug_Report.xlsx")
# kenya_ar=pd.read_excel("P:/Pertinent Files/Python/scripts/Finance/ky_Report.xlsx")
# customer_ug=pd.read_excel('P:/Pertinent Files/Python/scripts/Finance/customer_ug.xlsx')
# customer_ky=pd.read_excel('P:/Pertinent Files/Python/scripts/Finance/customer_ky.xlsx')



# print(uganda_ar.columns)

#%%

AR_Standing_status = pd.ExcelWriter('P:/Pertinent Files/Python/scripts/Finance/Customer Credit Report.xlsx')

# Write kenya_ar data to a sheet named "Kenya AR"
# #consolidated
# pd.concat([uganda_ar[['Description','CustomerPriorityClassificationGroupCode',f'As at: {date_ar} [ugx]']],
#            kenya_ar[['Description','CustomerPriorityClassificationGroupCode',f'As at: {date_ar} [kshs]']]]).to_excel(AR_Standing_status, sheet_name='Consolidated Summary By Sales Channel',index=False)


uganda_ar.to_excel(AR_Standing_status, sheet_name='By Sales Channel - UG',index=False)
customer_ug.to_excel(AR_Standing_status, sheet_name='Customer Details UG',index=False)

kenya_ar.to_excel(AR_Standing_status, sheet_name='By Sales Channel - KY',index=False)
customer_ky.to_excel(AR_Standing_status, sheet_name='Customer Details KY',index=False)


# Save the workbook
AR_Standing_status.close()


# #%%

# """
# ######################################################################
# # Email With Attachments Python Script

# ######################################################################
# """

# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from email.mime.base import MIMEBase
# from email import encoders

# # Setup port number and server name

# smtp_port = 587                 # Standard secure SMTP port
# smtp_server = "smtp-mail.outlook.com"  # Google SMTP Server

# # Set up the email lists
# email_from = pbi_user
# email_list = ["pokuttu@yalelo.ug"]
# # email_list = ["pokuttu@yalelo.ug","knyeko@yalelo.ug","rnabukeera@yalelo.ug","aoriide@yalelo.ug","alakica@yalelo.ug"]

# # Define the password (better to reference externally)
# pswd = pbi_pass # As shown in the video this password is now dead, left in as example only


# # name the email subject
# subject = f"Customer Credit Report {formatted_date}"

# import sys 
# new_path='P:/Pertinent Files/Python/scripts/Finance/'
# # sys.path.insert(0, new_path)
# os.chdir(new_path)
# # Define the email function (dont call it email!)
# def send_emails(email_list):

#     # Prepare the body of the email
#     body = """
#     Hello Team,
    
#     Please find Accounts Receivables Report attached.

#     Regards,
#     Audit Team
#     """

#     # Make a MIME object to define parts of the email
#     msg = MIMEMultipart()
#     msg['From'] = email_from
#     msg['Subject'] = subject
#     # recipients = '; '.join(email_list)
#      # Add CC recipients
#     # cc_recipients = email_list
#     # msg['CC'] = ', '.join(email_list)
#     msg['To'] = ', '.join(email_list)
#     # Attach the body of the message
#     msg.attach(MIMEText(body, 'plain'))

#     # Define the file to attach
#     filename = "Customer Credit Report.xlsx"

#     # Open the file in python as a binary
#     with open(filename, 'rb') as attachment:
#         # Encode as base 64
#         attachment_package = MIMEBase('application', 'octet-stream')
#         attachment_package.set_payload(attachment.read())
#         encoders.encode_base64(attachment_package)
#         attachment_package.add_header('Content-Disposition', f"attachment; filename= {filename}")
#         msg.attach(attachment_package)

#     # Collect all recipients
    

#     # Add CC recipients
#     # cc_recipients = ['poe@gmail.com', 'k@gmail.com']
#     # msg['CC'] = ', '.join(cc_recipients)

#     # Connect with the server
#     print("Connecting to server...")
#     TIE_server = smtplib.SMTP(smtp_server, smtp_port)
#     TIE_server.starttls()
#     TIE_server.login(email_from, pswd)
#     print("Succesfully connected to server")
#     print()

#     # Send emails to all recipients at once
#     print("Sending emails...")
#     TIE_server.sendmail(email_from,email_list, msg.as_string())
#     print("All emails sent")
#     print()

#     # Close the port
#     TIE_server.quit()

# # Run the function
# send_emails(email_list)
>>>>>>> c03a598df93b05e6595dc8f7aa81f3b56fa360c0
