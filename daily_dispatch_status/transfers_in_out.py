#%%
#import modules
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
#%%

from datetime import date,datetime, timedelta
import calendar
my_date = date.today()
#%%
#get email lists

active_warehouse = pd.read_csv('P:/Pertinent Files/Python/scripts/daily_dispatch_status/distinct_warehouse.csv')
active_warehouse.dropna(subset=['Email']).head()

#%%

day_of_the_week=calendar.day_name[my_date.weekday()]
day_of_the_week_num=datetime.today().weekday() #0 for Monday, 1 for Tuesday
month_start_date =my_date.replace(day=1).strftime('%m/%d/%Y')
month_end_date = my_date.strftime('%m/%d/%Y')


week_start_date=my_date - timedelta(days=my_date.weekday())
week_end_date=week_start_date+timedelta(days=6)

week_start_date=week_start_date.strftime('%m/%d/%Y')
week_end_date=week_end_date.strftime('%m/%d/%Y')

print(f"Month start date: {month_start_date}")
print(f"Month end date: {month_end_date}")

print(f"week start date: {week_start_date}")
print(f"week end date: {week_end_date}")
#%%
current_time = str(datetime.now().time())

# Options.add_experimental_option("detach", True)

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
# # send power bi login user
pbi_user='d365@yalelo.ke'
pbi_pass='Kenya@YK'
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

# Navigate to the report URL
driver.get(transfer_in_url)
# wait for report load

#Click to get Drop Down Menu

def export_tr(url):

    download_address=glob.glob("C:/Users/Pole Okuttu/Downloads/data"+ "*xlsx")

    file_path=[]

    for h in download_address:
        os.remove(h)

# load url
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

    # Get the latest downloaded file
    tr_file = pd.read_excel('C:/Users/Pole Okuttu/Downloads/data.xlsx')
    # Close the browser
    return tr_file

driver.quit()

#%%
#Read

transfers_in = export_tr(transfer_in_url)
transfers_out = export_tr(transfer_out_url)

#%%
# df_in = transfers_in.pivot(index=["Date"], columns=["ProductSizeId"], values="Received (kg)")
df_in = pd.pivot_table(transfers_in,index=["Date", "ReceivingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum').
# df_in['Total'] = df_in.sum(axis=1, skipna=True)
df_in

#%%

# df_in = transfers_in.pivot(index=["Date"], columns=["ProductSizeId"], values="Received (kg)")
df_out = pd.pivot_table(transfers_out,index=["Date","ReceivingWarehouseId","ShippingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum').reset_index
df_out['Total'] = df_out.sum(axis=1, skipna=True)
df_out.head()
# %%
df_in.merge(active_warehouse, on="WarehouseId", how="left")
# for j in active_warehouse['WarehouseId'].sort_values().to_list():
#     for i in transfers_in['ReceivingWarehouseId'].dropna().drop_duplicates().sort_values().to_list():
#         if i in 

# %%
