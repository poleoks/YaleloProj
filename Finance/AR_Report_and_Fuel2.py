import sys
save_dir = "C:/Users/Administrator/Documents/Python_Automations/"
# save_dir = "D:/YU/ScriptCodez/"
sys.path.append(f'{save_dir}')
from credentials import *   

for file in glob.glob("c:/Users/Administrator/Downloads/*.xlsx*"):
    os.remove(file)
    print(f"Deleted file: {file}")

for xl in glob.glob("C:/Users/Administrator/Documents/Python_Automations/Finance/*.xlsx*"):
    os.remove(xl)
    print(f"Deleted file: {xl}")


from powerbi_sign_in_file import *
from datetime import datetime#, timedelta


date_today=datetime.today().day
month_num=datetime.today().month
year=datetime.today().year
# Format the date as "mm/dd/yyyy" using strftime()
formatted_date = datetime.today().strftime('%m/%d/%Y')
date_ar= datetime.today().strftime('%Y-%m-%d')
print(formatted_date)
file_location='c:/Users/Administrator/Downloads/Customer aging report.xlsx'

#GET AR REPORT FOR KENYA
browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yk&mi=Output%3ACustAgingBalance')
browser.maximize_window()

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
    
#ok button to load report
browser.find_element(By.XPATH, '//*[@id="SysOperationTemplateForm_2_CommandButton"]').click()


#wait for load and press export button

search_xpath = '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportMenuButton_label"]'
search_button = WebDriverWait(browser, 500).until(
    EC.presence_of_element_located((By.XPATH, search_xpath))
)

search_button.click()
time.sleep(2)
#%%

WebDriverWait(browser, 500).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportToExcelButton_label"]'))).click()
time.sleep(5)


bb=f"As at: {date_ar} [kshs]"
try:
    file_location='c:/Users/Administrator/Downloads/Customer aging report.xlsx'
except:
    None
bb=f"As at: {date_ar} [kshs]"
df=pd.read_excel(file_location,skiprows=11,skipfooter=1)
df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
df.rename(columns={ df.columns[3]:bb,df.columns[4]:"Current (0-7) days [kshs]",df.columns[5]:"8-14 days [kshs]",df.columns[6]:"15-29 days [kshs]",
                df.columns[7]:"30-89 days [kshs]",df.columns[8]:"90+ [kshs]"}, inplace = True)
kenya_ar_extract=df.copy()
kenya_ar_extract.to_excel('C:/Users/Administrator/Documents/Python_Automations/Finance/Customer AR Kenya.xlsx',index=False)
#%%
#GET AR REPORT FOR UGANDA
for file in glob.glob("c:/Users/Administrator/Downloads/*.xlsx*"):
    os.remove(file)
    print(f"Deleted file: {file}")
time.sleep(5)
print("Kenya extraction complete, Starting Uganda AR extraction")

browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=Output%3ACustAgingBalance')
browser.maximize_window()
time.sleep(5)

aging_date=WebDriverWait(browser, 500).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="SysOperationTemplateForm_2_Fld3_1_input"]'))
)

aging_date.clear()
aging_date.send_keys(formatted_date)
time.sleep(2)

balance_date=browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld4_1_input"]')
balance_date=browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld4_1_input"]')
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
# browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_input"]').click()
time.sleep(2)

# browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld5_1_list_item0"]').click()
time.sleep(2)

#currency filter
currency_filter=browser.find_element(By.XPATH,'//*[@id="SysOperationTemplateForm_2_Fld7_1_input"]')
currency_filter.clear()

time.sleep(2)
currency_filter.send_keys("Accounting currency")
time.sleep(2)

#negative balance
try:
    browser.find_element(By.XPATH, "//*[@id='SysOperationTemplateForm_2_Fld15_1']//span[@class='toggle-value' and @title='No']").click()
    toggle_element = browser.find_element(By.XPATH, "//span[@id='SysOperationTemplateForm_2_Fld15_1_toggle']")

    browser.find_element(By.XPATH, "//*[@id='SysOperationTemplateForm_2_Fld15_1']//span[@class='toggle-value' and @title='No']").click()
    toggle_element = browser.find_element(By.XPATH, "//span[@id='SysOperationTemplateForm_2_Fld15_1_toggle']")
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

search_button = WebDriverWait(browser, 500).until(
    EC.presence_of_element_located((By.XPATH, search_xpath))
)

search_button.click()
time.sleep(2)

print("Export Menu Clicked")
browser.find_element(By.XPATH, '//*[@id="SrsReportPdfViewerForm_5_PdfViewerExportToExcelButton_label"]').click()
time.sleep(5)


# Wait for the file to download (this might need adjustments based on download time)
WebDriverWait(browser, 5*60).until(
    lambda driver: len(glob.glob(f"{file_location}")) > len(f_path)
)

print("Export to Excel completed")
time.sleep(2)
browser.quit()
# time.sleep(5)

bb=f"As at: {date_ar} [ugx]" 
df=pd.read_excel(file_location,skiprows=11,skipfooter=1)
df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
df.rename(columns={ df.columns[3]:bb,df.columns[4]:"90+ days [ugx]",df.columns[5]:"90 days [ugx]",df.columns[6]:"60 days [ugx]",
                df.columns[7]:"30 days [ugx]",df.columns[8]:"Current Date [ugx]"}, inplace = True)
uganda_ar_extract = df.copy()
uganda_ar_extract.to_excel(f"{save_dir}/Customer AR Uganda.xlsx",index=False)



#%%

#%%
AR_Standing_status = pd.ExcelWriter('C:/Users/Administrator/Documents/Python_Automations/Finance/Customer Credit Report.xlsx')


uganda_ar_extract.to_excel(AR_Standing_status, sheet_name='Customer Details UG',index=False)
kenya_ar_extract.to_excel(AR_Standing_status, sheet_name='Customer Details KY',index=False)

# Save the workbook
AR_Standing_status.close()


time.sleep(5)
#%%

"""
######################################################################
# Email With Attachments Python Script

######################################################################
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Setup port number and server name

smtp_port = 587                 # Standard secure SMTP port
smtp_server = "smtp-mail.outlook.com"  # Google SMTP Server

# Set up the email lists
email_from = pbi_user
email_list = ["pokuttu@yalelo.ug"]
# email_list = ["pokuttu@yalelo.ug","knyeko@yalelo.ug","rnabukeera@yalelo.ug","aoriide@yalelo.ug","alakica@yalelo.ug"]

# Define the password (better to reference externally)
pswd = pbi_pass # As shown in the video this password is now dead, left in as example only


# name the email subject
subject = f"Customer Credit Report - AR"

# import sys 
new_path='C:/Users/Administrator/Documents/Python_Automations/Finance/'
# sys.path.insert(0, new_path)
os.chdir(new_path)
# Define the email function (dont call it email!)
body = """
Hello Team,

Please find Accounts Receivables Report attached.

Regards,
Audit Team
"""

# Define the file to attach
filename = "Customer Credit Report.xlsx"

from gmail_sender import gmail_function

gmail_function('pokuttu@yalelo.ug', subject_line=subject, body=body, filename= filename)

try:
    os.remove(filename)
    print(f"{filename} file deleted")
except:
    print("no file found")
    pass

#%%
time.sleep(60)
kill_browser("chrome")