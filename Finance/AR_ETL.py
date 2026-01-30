import sys
save_dir = "C:/Users/Administrator/Documents/Python_Automations/"
# save_dir = "D:/YU/ScriptCodez/"
sys.path.append(f'{save_dir}')
from credentials import *   

for file in glob.glob("c:/Users/Administrator/Downloads/*.xlsx*"):
    os.remove(file)
    print(f"Deleted file: {file}")

for i in glob.glob("C:/Users/Administrator/Documents/Python_Automations/Finance/"+"*xlsx"):
        os.remove(i)
        print(f"{i} removed")   

from powerbi_sign_in_file import *
from datetime import datetime#, timedelta


date_today=datetime.today().day
month_num=datetime.today().month
year=datetime.today().year
# Format the date as "mm/dd/yyyy" using strftime()
formatted_date = datetime.today().strftime('%m/%d/%Y')
date_ar= datetime.today().strftime('%Y-%m-%d')
print(formatted_date)
file_location=f"Customer Credit Report - {date_ar}.xlsx"
download_path="C:/Users/Administrator/Downloads"
download_address = "C:/Users/Administrator/Downloads/Customer aging report.xlsx"

#GET AR REPORT FOR KENYA
time.sleep(5)
browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yk&mi=Output%3ACustAgingBalance')
time.sleep(15)
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


WebDriverWait(browser, 1*60).until(
    lambda driver: len(glob.glob(f"{download_address}")) > len(f_path)
)
print("Export to Excel completed")
time.sleep(2)


bb=f"As at: {date_ar} [kshs]"
df=pd.read_excel(download_address,skiprows=11,skipfooter=1)

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
    lambda driver: len(glob.glob(f"{download_address}")) > len(f_path)
)

print("Export to Excel completed")
time.sleep(2)
browser.quit()
# time.sleep(5)

bb=f"As at: {date_ar} [ugx]" 
df=pd.read_excel(download_address,skiprows=11,skipfooter=1)
df=df[df.columns.drop(list(df.filter(regex='Unnamed')))]
df.rename(columns={ df.columns[3]:bb,df.columns[4]:"90+ days [ugx]",df.columns[5]:"90 days [ugx]",df.columns[6]:"60 days [ugx]",
                df.columns[7]:"30 days [ugx]",df.columns[8]:"Current Date [ugx]"}, inplace = True)
uganda_ar_extract = df.copy()
uganda_ar_extract.to_excel(f"{save_dir}Finance/Customer AR Uganda.xlsx",index=False)

#%%
AR_Standing_status = pd.ExcelWriter(file_location, engine='xlsxwriter')

uganda_ar_extract.to_excel(AR_Standing_status, sheet_name=f'Customer Details UG',index=False)
kenya_ar_extract.to_excel(AR_Standing_status, sheet_name=f'Customer Details KY',index=False)

# Save the workbook
AR_Standing_status.close()
time.sleep(5)
#%%
# name the email subject
subject = f"Customer Credit Report - AR"

# Define the email function (dont call it email!)
body = "Hello Team,\nPlease find Accounts Receivables Report attached.\n\nRegards,\nAudit Team"
# Define the file to attach


from gmail_sender import gmail_function

gmail_function('pokuttu@yalelo.ug', subject_line=subject, email_body=body, attachment_path=file_location)
try:
    for i in glob.glob("C:/Users/Administrator/Documents/Python_Automations/Finance/"+"*xlsx"):
        os.remove(i)
        print(f"{i} removed")
except:
    print("no file found")
    pass
time.sleep(60)
kill_browser("chrome")