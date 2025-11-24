import sys
#file path
# file_path_r = 'C:/Users/Administrator/Documents/Python_Automations/'
# download_path = 'C:/Users/Administrator/Downloads/'

file_path_r = 'D:/YU/ScriptCodez/'
download_path = 'C:/Users/Pole Okuttu/Downloads/'

sys.path.append(file_path_r)
from powerbi_sign_in_file import *
print("Modules imported successfully")

today = datetime.datetime.today()
first_day = today.replace(day=12) - relativedelta(months=0)

# Calculate the last day of the month
last_day = (first_day + relativedelta(months=1)) - timedelta(days=1)
first_day=today.strftime('%m/%d/%Y')
last_day=last_day.strftime('%m/%d/%Y')
# print(range(12))
print(f"First day: {first_day}, Last day: {last_day}")
#%%
browser.delete_all_cookies()
time.sleep(2)
browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=InventTransNew')
#%%
time.sleep(10)
gridopt=WebDriverWait(browser,45).until(
                EC.presence_of_element_located((By.XPATH,'//*[@aria-label="Grid options"]'))
                )
print("click!")
gridopt.click()
print("grid clicked")


clk=WebDriverWait(browser,15).until(
                EC.presence_of_element_located((By.XPATH,'//*[@aria-label="Insert columns..."]'))
                )
print("click!")
clk.click()
print("insert cols")
time.sleep(2)
for cols in ["Warehouse"]:     
        filter = WebDriverWait(browser,5).until(
                        EC.presence_of_element_located((By.XPATH,'(//input[@aria-label="Filter"])[2]'))
                        )
        filter.clear()
        print("filter cleared")
        time.sleep(1)

        filter.send_keys(cols)
        print(f"{cols} pasted")

        time.sleep(2)
        browser.find_element(By.XPATH,'//*[@class="quickFilter-listFieldName" and contains(text(), "Field")]').click()
        time.sleep(2)

        browser.find_element(By.XPATH,'//*[contains(@class,"dyn-checkbox-span")]').click()
        time.sleep(2)
        print(f"New column:{cols} added")
cupd = WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,"//*[contains(text(),'Update')]"))
        )
cupd.click()
time.sleep(25)

# Add filters
st_date = WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,'(//*[@aria-label="Filter field: Financial date, operator: between"])[1]'))
        )

st_date.clear()
time.sleep(1)
st_date.send_keys(f"{first_day}")
print("filter cleared")
en_date = WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,'(//*[@aria-label="Filter field: Financial date, operator: between"])[2]'))
        )

en_date.clear()
time.sleep(1)
en_date.send_keys(f"{last_day}")


#%%
# Add custom filter Receipt
addf=WebDriverWait(browser,45).until(
                EC.presence_of_element_located((By.XPATH,'//*[@data-dyn-role="Button" and @class="filterDisplay-addFieldControl button dynamicsButton" ]'))
                )
addf.click()
print("---click add filter")
for cols in ["Receipt"]:     
        filter = WebDriverWait(browser,5).until(
                        EC.presence_of_element_located((By.XPATH,'(//input[@aria-label="Filter"])[2]'))
                        )
        filter.clear()
        print("filter cleared")
        time.sleep(1)

        filter.send_keys(cols)
        print(f"{cols} pasted")

        time.sleep(2)
        browser.find_element(By.XPATH,'//*[@class="quickFilter-listFieldName" and contains(text(), "Field")]').click()
        time.sleep(2)

        browser.find_element(By.XPATH,'//*[contains(@class,"dyn-checkbox-span")]').click()
        time.sleep(2)
        print(f"New filter:{cols} added")
cupd = WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,"//*[contains(text(),'Update')]"))
        )
cupd.click()
time.sleep(5)
#%%
#load report
##Receitp filter to Purchased only
rct=WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,'//*[@aria-label="Filter field: Receipt status (Receipt), operator: is exactly"]'))
        )
rct.send_keys('Purchased')
print("Receipt filtered to Purchased only")

time.sleep(2)
#click apply filters
WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@data-dyn-controlname="SystemDefinedFilterPane_FilterDisplay_ApplyFilters"]'))
                ).click()
time.sleep(15)

#click microsoft icon
WebDriverWait(browser, 60).until(
    EC.presence_of_element_located(
        (By.XPATH,'//*[@class="button-commandRing MicrosoftOffice-symbol"]')
    )
).click()

#click export excel
WebDriverWait(browser, 60).until(
    EC.presence_of_element_located(
        (By.XPATH,'//*[@aria-label="Export to Excel Inventory transactions originator"]')
    )
).click()
file_path =[]
#click download button
WebDriverWait(browser, 60).until(
    EC.presence_of_element_located(
        (By.XPATH,'//*[@data-dyn-controlname="DownloadButton"]')
    )
).click()
print("Downloading Inventory transactions excel file...")
WebDriverWait(browser, 30*60).until(
                    lambda driver: len(glob.glob(f"{download_path}Inventory transactions*.xlsx")) > len(file_path)
                    )


for file in glob.glob(f"{download_path}Inventory transactions*.xlsx"):
                    dd = pd.read_excel(file)
                    print(dd.head())
dd.to_excel('Inventory Transactions.xlsx', index=False)
time.sleep(5)
print(dd.head())
browser.quit()
#%%
from gmail_sender import *
subject = "Inventory Transactions"
body = "Hi Team,\n\nPlease find attached the Inventory Transactions report.\n\nBest regards,\nPole"
to = "pokuttu@yalelo.ug"
attachments = ['Inventory Transactions.xlsx']
gmail_function(to,subject, body, attachments)

#%%
#cleanup downloads folder
time.sleep(5)
for i in glob.glob(f"{download_path}Inventory transactions*.xlsx"):
    os.remove(i)
    print(f"{i} removed")
