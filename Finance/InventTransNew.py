import sys
#file path
file_path_r = 'C:/Users/Administrator/Documents/Python_Automations/'
download_path = 'C:/Users/Administrator/Downloads/'

# file_path_r = 'D:/YU/ScriptCodez/'
# download_path = 'C:/Users/Pole Okuttu/Downloads/'

sys.path.append(file_path_r)
from powerbi_sign_in_file import *
print("Modules imported successfully")

today = datetime.datetime.today()
# set today as yesterday
today = today - timedelta(days=1)
year_ = today.year
month_ = today.month
# end_day_filter = today.replace(day=12) - relativedelta(months=0)

# # Calculate the first day of the month
start_day_filter = today.replace(day=1)
end_day_filter=today.strftime('%m/%d/%Y')
start_day_filter=start_day_filter.strftime('%m/%d/%Y')
attachments_ = f'Inventory Transactions-{year_}-{month_}.xlsx'
print(f"First day: {end_day_filter}, Last day: {start_day_filter}")
#%%
#cleanup downloads folder
tttt=file_path_r+attachments_

for i in glob.glob(f"{download_path}Inventory transactions*.xlsx"):
    try:
            os.remove(i)
            print(f"{i} removed")
    except FileNotFoundError:
            print(f"{i} not found, skipping removal.")

try:
        os.remove(tttt)
except FileNotFoundError:
        print(f"{tttt} not found, skipping removal.")       
            
print(f"{tttt} removed")

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
        filter = WebDriverWait(browser,55).until(
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


#%%
# Add custom filter Receipt
addf=WebDriverWait(browser,45).until(
                EC.presence_of_element_located((By.XPATH,'//*[@data-dyn-role="Button" and @class="filterDisplay-addFieldControl button dynamicsButton" ]'))
                )
addf.click()
print("---click add filter")
for cols in ["Receipt"]:     
        filter = WebDriverWait(browser,55).until(
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
# Add filters
st_date = WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,'(//*[@aria-label="Filter field: Financial date, operator: between"])[1]'))
        )

st_date.clear()
time.sleep(1)
st_date.send_keys(f"{start_day_filter}")
print("filter cleared")
en_date = WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,'(//*[@aria-label="Filter field: Financial date, operator: between"])[2]'))
        )

en_date.clear()
time.sleep(1)
en_date.send_keys(f"{end_day_filter}")


##Receitp filter to Purchased only
# rct=WebDriverWait(browser,60).until(
#         EC.presence_of_element_located((By.XPATH,'//*[@aria-label="Filter field: Receipt status (Receipt), operator: is exactly"]'))
#         )
# rct.clear()
# time.sleep(1)
# rct.send_keys('Purchased')
# print("Receipt filtered to Purchased only")

fnb=WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,'//*[@aria-label="Filter field: Number, operator: begins with"]'))
        )
fnb.clear()
time.sleep(1)
fnb.send_keys('YUTO')
print("Number filtered to start with YUTO only")

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
time.sleep(5)
print("Downloading Inventory transactions excel file...")
WebDriverWait(browser, 30*60).until(
                    lambda driver: len(glob.glob(f"{download_path}/Inventory transactions*.xlsx")) > len(file_path)
                    )
print("Download completed!!!")

for file in glob.glob(f"{download_path}Inventory transactions*.xlsx"):
                    dd = pd.read_excel(file)
                    print(dd.head())
dd.to_excel(attachments_, index=False)

print(dd.head())
time.sleep(5)
browser.quit()
#%%
from gmail_sender import *
subject = "Inventory Transactions YTD"
body = "Hi Team,\n\nPlease find attached the Inventory Transactions report.\n\nBest regards,\nPole"
to = "pokuttu@yalelo.ug"
gmail_function(to,subject, body, attachments_)
time.sleep(5)
#%%
#cleanup downloads folder
tttt=file_path_r+attachments_


for i in glob.glob(f"{download_path}Inventory transactions*.xlsx"):
    try:
            os.remove(i)
            print(f"{i} removed")
    except FileNotFoundError:
            print(f"{i} not found, skipping removal.")

try:
        os.remove(tttt)
except FileNotFoundError:
        print(f"{tttt} not found, skipping removal.")       
            
print(f"{tttt} removed")