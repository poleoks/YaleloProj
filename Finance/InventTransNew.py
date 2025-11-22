import sys
#file path
file_path_r = 'C:/Users/Administrator/Documents/Python_Automations/'
download_path = 'C:/Users/Administrator/Downloads/'
sys.path.append(file_path_r)
from powerbi_sign_in_file import *
from datetime import datetime#, timedelta
#import calendar


today = datetime.today()
first_day = today.replace(day=1) - relativedelta(months=0)

# Calculate the last day of the month
last_day = (first_day + relativedelta(months=12)) - timedelta(days=1)
first_day=today.strftime('%m/%d/%Y')
last_day=last_day.strftime('%m/%d/%Y')
# print(range(12))

#%%

browser.delete_all_cookies()
browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=InventTransNew')

#load report
##Reference
WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,'//*[@data-dyn-columnname="InventTransOrigin_ReferenceCategory"]'))
        ).click()
print("Reference")
bt=WebDriverWait(browser,60).until(
        EC.presence_of_element_located((By.XPATH,'//*[@aria-label="Filter field: Reference, operator: is exactly"]'))
        )
print("send key = Sales Orders")
bt.clear()
time.sleep(1)
bt.send_keys('Sales order')

time.sleep(3)


WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@data-dyn-controlname="SystemDefinedFilterPane_FilterDisplay_ApplyFilters"]'))
                ).click()
time.sleep(5)

WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@data-dyn-qtip-title="Financial date"]'))
                ).click()

WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@class="filterDisplay-addFieldControl button dynamicsButton"]'))
                ).click()
print("send key = Sales Orders ll")
sk = WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@name="QuickFilterControl_Input"]'))
                )
sk.send_keys("Financial Date")

chk_b = WebDriverWait(browser,60).until(
    EC.presence_of_element_located((By.XPATH,'//*[@type="checkbox"]'))
    )[0]

chk_b.click()


st_date = WebDriverWait(browser,60).until(
    EC.presence_of_element_located((By.XPATH,'//*[@name="FilterField_InventTrans_DateFinancial_DateFinancial_Input_0"]'))
    )
st_date.clear()
time.sleep(1)
st_date.send_keys(first_day)


en_date = WebDriverWait(browser,60).until(
    EC.presence_of_element_located((By.XPATH,'//*[@name="FilterField_InventTrans_DateFinancial_DateFinancial_Input_1"]'))
    )

en_date.clear()
time.sleep(1)
en_date.send_keys(last_day)


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

time.sleep(5)
for i in glob.glob(f"{download_path}Inventory transactions*.xlsx"):
    os.remove(i)
    print(f"{i} removed")
