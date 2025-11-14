#%%
import glob
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

from datetime import date,datetime, timedelta
import calendar
my_date = date.today()
#get email lists
#%%
active_warehouse = [
    {'WarehouseId' : 'Bulaga', 'Email':'bulagaduuka@yalelo.ug'},
    {'WarehouseId' : 'Bunamwaya', 'Email':'bunamwayastore@yalelo.ug'},
    {'WarehouseId' : 'Busia', 'Email':'busiastore@yalelo.ug'},
    {'WarehouseId' : 'Bwaise', 'Email':'bwaisestore@yalelo.ug'},
    {'WarehouseId' : 'Distrbtion', 'Email':'mbiryetega@yalelo.ug'},
    {'WarehouseId' : 'Jinja V3', 'Email':'jinjastore@yalelo.ug'},
    {'WarehouseId' : 'Kafunta', 'Email':'kafuntastore@yalelo.ug'},
    {'WarehouseId' : 'Kasangati', 'Email':'kasangatistore@yalelo.ug'},
    {'WarehouseId' : 'Kasubi', 'Email':'kasubistore@yalelo.ug'},
    {'WarehouseId' : 'Kawempe V2', 'Email':'kawempestore@yalelo.ug'},
    {'WarehouseId' : 'Kibuli', 'Email':'kibulistore@yalelo.ug'},
    {'WarehouseId' : 'Kibuye', 'Email':'kibuyestore@yalelo.ug'},
    {'WarehouseId' : 'Kireka', 'Email':'kirekastore@yalelo.ug'},
    {'WarehouseId' : 'Kyaliwajal', 'Email':'kyaliwajjalastore@yalelo.ug'},
    {'WarehouseId' : 'Kyambogo', 'Email':'kyambogostore@yalelo.ug'},
    {'WarehouseId' : 'Kyengera-R', 'Email':'kyengerastore@yalelo.ug'},
    {'WarehouseId' : 'Malaba', 'Email':'malababorder@yalelo.ug'},
    {'WarehouseId' : 'Mpondwe', 'Email':'mpondweborder@yalelo.ug'},
    {'WarehouseId' : 'Mukono', 'Email':'mukonoshop@yalelo.ug'},
    {'WarehouseId' : 'Mutungo', 'Email':'mutungostore@yalelo.ug'},
    {'WarehouseId' : 'Nansana V2', 'Email':'nansanastore@yalelo.ug'},
    {'WarehouseId' : 'Natete V2', 'Email':'nateetestore@yalelo.ug'},
    {'WarehouseId' : 'Ntinda', 'Email':'ntindastore@yalelo.ug'},
    {'WarehouseId' : 'Nyahuka', 'Email':'nyahukaborder@yalelo.ug'}
    ]


active_warehouse = pd.DataFrame(active_warehouse)

#%%

Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)

driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
# powerbi login url
driver.get('https://app.powerbi.com/?noSignUpCheck=1')
driver.maximize_window()


# wait for page load
time.sleep(5)
# find email area using xpath
email=WebDriverWait(driver, 5*60).until(
    EC.presence_of_element_located((By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']"))
)
#%%
# # send power bi login user
pbi_user='poweruser@firstwave.ag'
pbi_pass='Vaj37993'

#%%
email.send_keys(pbi_user)

#find submission data
submit = driver.find_element("id",'idSIButton9')
# # click for submit button
submit.click()

time.sleep(5)
# now we need to find password field
password = WebDriverWait(driver, 5*60).until(
    EC.presence_of_element_located(("id",'i0118'))
)
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
#navigate the report and take screenshot
current_dir= "P:/Pertinent Files/Python/scripts/daily_dispatch_status/"
ss_name ='distribution.png'
ss_path=current_dir+ss_name
# ss_path = os.path.join(current_dir, "/", ss_name)
# print(f"This is the picture path: {ss_path}")

transfer_in_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/e4e07d5f73507b70b7bd"
transfer_out_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/203e64e3a7c660b4d76a"

download_address=glob.glob("C:/Users/Pole Okuttu/Downloads/data"+"*xlsx")

file_path=[]
    # load url
def export(url):

    for h in download_address:
        os.remove(h)

    driver.get(url)
    hover_element=WebDriverWait(driver, 5*60).until(
        EC.presence_of_element_located((By.XPATH,'//*[ @role="presentation" and @class="top-viewport"]'))
    )

    ActionChains(driver).double_click(hover_element).perform()

    WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH,'//*[@data-testid="visual-more-options-btn" and @class="vcMenuBtn" and @aria-expanded="false" and @aria-label="More options"]'))
        ).click()

    #Click Export data
    WebDriverWait(driver, 3*60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@title="Export data"]'))
    ).click()

    #Click Export Icon
    WebDriverWait(driver, 3*60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@aria-label="Export"]'))
    ).click()

    # Wait for the file to download (this might need adjustments based on download time)
    WebDriverWait(driver, 3*60).until(
        lambda driver: len(glob.glob("C:/Users/Pole Okuttu/Downloads/data*.xlsx")) > len(file_path)
    )
    return print("file download completed!!!")

#%%
# read files
export(transfer_in_url)

#%%
tr_in = pd.read_excel("C:/Users/Pole Okuttu/Downloads/data.xlsx")

#%%
df_in = pd.pivot_table(tr_in,index=["Date", "ReceivingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])
df_in['Total'] = df_in.sum(axis=1, skipna=True)


# #%%
# #DC transfers from production
# df_in_dc = pd.pivot_table(tr_in[tr_in['Received (kg)'] > 15],index=["Date", "ReceivingWarehouseId","ShippingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])
# df_in_dc['Total Received'] = df_in_dc.sum(axis=1, skipna=True)

# print(df_in_dc.head())
# #%%
# #DC driploss - from shops
# df_dl_dc = pd.pivot_table(tr_in[(tr_in['ShippingWarehouseId'] != 'Production') & (tr_in['Received (kg)'] < 15)],index=["Date", "ReceivingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])
# df_dl_dc['Total Driploss From Shops'] = df_dl_dc.sum(axis=1, skipna=True)

# print(df_dl_dc.head())
# #%%
# #pivot transfers_out
# export(transfer_out_url)
# tr_out = pd.read_excel("C:/Users/Pole Okuttu/Downloads/data.xlsx")
# df_out = pd.pivot_table(tr_out,index=["Date","ReceivingWarehouseId","ShippingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index()
# df_out['Total Shipped Out'] = df_out.sum(axis=1, skipna=True)

# #%%


# #%%

# driver.quit()
# #%%
# print(df_in.columns)

# #%%
# print(df_out.head())

# #%%
# # log-in to email server
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

# # Define the password (better to reference externally)
# pswd = pbi_pass # As shown in the video this password is now dead, left in as example only

# print("Connecting to server...")
# TIE_server = smtplib.SMTP(smtp_server, smtp_port)
# TIE_server.starttls()
# TIE_server.login(email_from, pswd)
# print("Succesfully connected to server")
# #%%
# for i,e in zip(active_warehouse['WarehouseId'].to_list(), active_warehouse['Email'].to_list()):
#     email_list = [f"{e}","pokuttu@yalelo.ug"]
#     try:
#         df_in_w = df_in.loc[(slice(None), i),:].dropna(axis=1, how='all')
#         df_in_w.loc['Sub Total', :] = df_in_w.sum().values
#     except:
#         None
#     try:
#         df_out_w = df_out.loc[slice(None), slice(None), i].dropna(axis=1, how='all')
#         df_out_w.loc['Sub Total', :] = df_out_w.sum().values
#     except:
#         None
#     if i=='Distrbtion':
#         with pd.ExcelWriter(f"{i} MTD Stock.xlsx") as writer:
#                 # Write df_in_w to the first sheet
#                 df_in_dc.to_excel(writer, sheet_name=f'{i} Stock Received From Production')
                               
#                 # Write df_in_w to the first sheet
#                 df_in_dc.to_excel(writer, sheet_name=f'{i} Stock Received From Production')

#                 # Write df_out_w to the second sheet
#                 df_out_w.to_excel(writer, sheet_name=f'{i} Stock Shipped Out')

#     else:
#         with pd.ExcelWriter(f"{i} MTD Stock.xlsx") as writer:
#             # Write df_in_w to the first sheet
#             df_in_w.to_excel(writer, sheet_name=f'{i} Stock Received')

#             # Write df_out_w to the second sheet
#             df_out_w.to_excel(writer, sheet_name=f'{i} Stock Shipped')

    
#    # name the email subject

#     subject = f"{i} Transfers Report [Month To Date]"

#     # import sys 
#     new_path='P:/Pertinent Files/Python/scripts/daily_dispatch_status/'

#     # Prepare the body of the email
#     body = f"""
#     Hello {i} Store,
    
#     Please find your month to date transfers in and out report attached. 
    
#     Check both sheets to ensure that you are aware and own this records.
#     The Total Received - Total Shipped Stock = Net Received (shared daily on whatsapp)
    
#     In case of any issues, please reply to pokuttu@yalelo.ug for followup.

#     Regards,
#     Pole
#     """
#     # Make a MIME object to define parts of the email
#     msg = MIMEMultipart()
#     msg['From'] = email_from
#     msg['Subject'] = subject
#     msg['To'] = ', '.join(email_list)


#     # Attach the body of the message
#     msg.attach(MIMEText(body, 'plain'))

#     # Define the file to attach
#     filename = f"{i} MTD Stock.xlsx"

#     # Open the file in python as a binary
#     with open(filename, 'rb') as attachment:
#         # Encode as base 64
#         attachment_package = MIMEBase('application', 'octet-stream')
#         attachment_package.set_payload(attachment.read())
#         encoders.encode_base64(attachment_package)
#         attachment_package.add_header('Content-Disposition', f"attachment; filename= {filename}")
#         msg.attach(attachment_package)
#     # Send emails to all recipients at once
#     TIE_server.sendmail(email_from,email_list, msg.as_string())
#     print(f"{i} email sent!")
# #%%
# print("All Emails Sent successfully")
# #clear space
# TIE_server.quit()
# #%%
# for k in glob.glob("P:/Pertinent Files/Python/scripts/daily_dispatch_status/"+ "*MTD Stock*" + "*.xlsx"):
#     os.remove(k)

# print("All files removed from repository")
# # %%

