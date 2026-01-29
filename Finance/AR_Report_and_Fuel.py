import sys
save_dir = "C:/Users/Administrator/Documents/Python_Automations/"
sys.path.append(f'{save_dir}')
# sys.path.append('D:/YU/ScriptCodez/')
from powerbi_sign_in_file import *
from datetime import datetime#, timedelta

try:
    for file in glob.glob(os.path.join("c:/Users/Administrator/Downloads","*.xlsx*")):
        os.remove(file)
    for xl in glob.glob(os.path.join("C:/Users/Administrator/Documents/Python_Automations/Finance/","*.xlsx*")):
        os.remove(xl)
    print("File deleted")
except:
    print("File Does not Exist")

date_today=datetime.today().day
month_num=datetime.today().month
year=datetime.today().year
# Format the date as "mm/dd/yyyy" using strftime()
formatted_date = datetime.today().strftime('%m/%d/%Y')
date_ar= datetime.today().strftime('%Y-%m-%d')
print(formatted_date)


#GET AR REPORT FOR KENYA
browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yk&mi=Output%3ACustAgingBalance')
browser.maximize_window()

# aging_date=WebDriverWait(browser, 500).until(
#     EC.presence_of_element_located((By.XPATH, '//*[@id="SysOperationTemplateForm_2_Fld3_1_input"]'))
# )
# aging_date.clear()
# aging_date.send_keys(formatted_date)
# time.sleep(2)

# balance_date=browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld4_1_input"]')
# balance_date.clear()
# balance_date.send_keys(formatted_date)
# time.sleep(2)

# # criteria filter selection
# browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_input"]').click()
# time.sleep(2)

# browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_list_item0"]').click()
# time.sleep(2)


# #aging period definition
# browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld6_1"]/div/div').click()
# time.sleep(2)

# browser.find_element(By.XPATH,'//input[@value="YKCustomerAging"]').click()
# time.sleep(2)

# #currency filter
# browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld7_1"]/div/div[2]/div').click()
# time.sleep(2)

# browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld7_1_list_item0"]').click()
# time.sleep(2)

# #negative balance
# try:
#     browser.find_element(By.XPATH, "//*[@id='SysOperationTemplateForm_2_Fld15_1']//span[@class='toggle-value' and @title='No']").click()
#     toggle_element = browser.find_element(By.XPATH, "//span[@id='SysOperationTemplateForm_2_Fld15_1_toggle']")
#     toggle_element.click()  # Simulate a click
# except:
#     print("switched on")
    
# #ok button to load report
# browser.find_element(By.XPATH, '//*[@id="SysOperationTemplateForm_2_CommandButton"]').click()


# #wait for load and press export button

# search_xpath = '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportMenuButton_label"]'
# search_button = WebDriverWait(browser, 500).until(
#     EC.presence_of_element_located((By.XPATH, search_xpath))
# )

# search_button.click()
# time.sleep(2)
# #%%

# WebDriverWait(browser, 500).until(
#     EC.presence_of_element_located((By.XPATH, '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportToExcelButton_label"]'))).click()
# time.sleep(5)

# browser.quit()


# bb=f"As at: {date_ar} [kshs]"
# try:
#     file_location='c:/Users/Administrator/Downloads/Customer aging report.xlsx'
# except:
#     None
# bb=f"As at: {date_ar} [kshs]"
# df=pd.read_excel(file_location,skiprows=11,skipfooter=1)
# df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
# df.rename(columns={ df.columns[3]:bb,df.columns[4]:"Current (0-7) days [kshs]",df.columns[5]:"8-14 days [kshs]",df.columns[6]:"15-29 days [kshs]",
#                 df.columns[7]:"30-89 days [kshs]",df.columns[8]:"90+ [kshs]"}, inplace = True)
# kenya_ar_extract=df.copy()
# kenya_ar_extract.to_excel('C:/Users/Administrator/Documents/Python_Automations/Finance/Customer AR Kenya.xlsx',index=False)
#%%
#GET AR REPORT FOR UGANDA
try:
    for x in glob.glob(os.path.join("c:/Users/Administrator/Downloads","Customer aging report*")):
        os.remove(x)
except:
    print('does not exist')

#%%
browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=Output%3ACustAgingBalance')
browser.maximize_window()

#%%
browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=Output%3ACustAgingBalance')
browser.maximize_window()
time.sleep(5)
drvr = browser

aging_date=WebDriverWait(drvr, 500).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="SysOperationTemplateForm_2_Fld3_1_input"]'))
)
aging_date.clear()
aging_date.send_keys(formatted_date)
time.sleep(2)

balance_date=browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld4_1_input"]')
balance_date=drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld4_1_input"]')
balance_date.clear()
balance_date.send_keys(formatted_date)
time.sleep(2)

# criteria filter selection
browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_input"]').click()
time.sleep(2)

browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_list_item0"]').click()
time.sleep(2)

#currency filter
# currency_filter=browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld7_1_input"]')
time.sleep(2)
# drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_input"]').click()
time.sleep(2)

# drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_list_item0"]').click()
time.sleep(2)

#currency filter
currency_filter=drvr.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld7_1_input"]')
currency_filter.clear()

time.sleep(2)
currency_filter.send_keys("Accounting currency")
time.sleep(2)

#negative balance
try:
    browser.find_element(By.XPATH, "//*[@id='SysOperationTemplateForm_2_Fld15_1']//span[@class='toggle-value' and @title='No']").click()
    toggle_element = browser.find_element(By.XPATH, "//span[@id='SysOperationTemplateForm_2_Fld15_1_toggle']")

    drvr.find_element(By.XPATH, "//*[@id='SysOperationTemplateForm_2_Fld15_1']//span[@class='toggle-value' and @title='No']").click()
    toggle_element = drvr.find_element(By.XPATH, "//span[@id='SysOperationTemplateForm_2_Fld15_1_toggle']")
    toggle_element.click()  # Simulate a click
except:
    print("switched on")

#aging period definition
browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld6_1"]/div/div').click()
time.sleep(2)

# browser.find_element(By.XPATH,'//input[@value="AgingYU"]').click()
time.sleep(2)
# ok button to load report
browser.find_element(By.XPATH, '//*[@id="SysOperationTemplateForm_2_CommandButton"]').click()

#wait for load and press export button

search_xpath = '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportMenuButton_label"]'

search_button = WebDriverWait(drvr, 500).until(
    EC.presence_of_element_located((By.XPATH, search_xpath))
)

search_button.click()
time.sleep(2)


browser.find_element(By.XPATH, '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportToExcelButton_label"]').click()
time.sleep(5)

browser.quit()
# time.sleep(5)

bb=f"As at: {date_ar} [ugx]"
df=pd.read_excel(save_dir,skiprows=11,skipfooter=1)
df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
df.rename(columns={ df.columns[3]:bb,df.columns[4]:"90+ days [ugx]",df.columns[5]:"90 days [ugx]",df.columns[6]:"60 days [ugx]",
                df.columns[7]:"30 days [ugx]",df.columns[8]:"Current Date [ugx]"}, inplace = True)

uganda_ar_extract = df.copy()
uganda_ar_extract.to_excel('C:/Users/Administrator/Documents/Python_Automations/Finance/Customer AR Uganda.xlsx',index=False)






#%%

# #%%
# AR_Standing_status = pd.ExcelWriter('C:/Users/Administrator/Documents/Python_Automations/Finance/Customer Credit Report.xlsx')


# uganda_ar_extract.to_excel(AR_Standing_status, sheet_name='Customer Details UG',index=False)
# kenya_ar_extract.to_excel(AR_Standing_status, sheet_name='Customer Details KY',index=False)
# ke_fuel.to_excel(AR_Standing_status, sheet_name='Kenya Fuel',index=False)
# ug_fuel.to_excel(AR_Standing_status, sheet_name='Uganda Fuel',index=False)

# # Save the workbook
# AR_Standing_status.close()


# time.sleep(5)
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
# subject = f"Customer Credit Report"

# # import sys 
# new_path='C:/Users/Administrator/Documents/Python_Automations/Finance/'
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

#     # Connect with the server
#     print("Connecting to server...")
#     TIE_server = smtplib.SMTP(smtp_server, smtp_port)
#     TIE_server.starttls()

#     TIE_server.login(email_from, pbi_pass_email)

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

# send_emails(email_list)
