save_dir= "C:/Users/Administrator/Documents/Python_Automations/"
import sys
sys.path.append(save_dir)
from credentials import *
from powerbi_sign_in_file import *
from warehouse_email import *
import glob
import time
import os
import pandas as pd
#%%
ebee_sales = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/91687127-028f-4c1e-8ad9-3f783f724150/562bc6b02aee23bc1cec"


#%%
download_address=glob.glob("c:/Users/Administrator/Downloads/data"+"*xlsx") #c:/Users\Administrator\Downloads
download_file_path = "C:/Users/Administrator/Downloads/data.xlsx"
file_path=[]
#%%
for file in download_address:
    try:
        os.remove(file)
        print(f"{file} removed")
    except:
        print("no file found")
# export pbi reports
ebee_rpt = pbi_export(ebee_sales,download_file_path)

browser.quit()

#%%
with pd.ExcelWriter('ebee_sales.xlsx') as wr:
    ebee_rpt.to_excel(wr, sheet_name="ebee_sales")
from gmail_sender import gmail_function

gmail_function('pokuttu@yalelo.ug, vguzman@yalelo.ug, btumwebaze@yalelo.ug, modern@ebee.africa, diana@ebee.africa,mathias@ebee.africa','ebee_sales','This is yesterday''s sales all ebee sales per location',
               'ebee_sales.xlsx')

try:
    os.remove('ebee_sales.xlsx')
    print("ebee_sales file deleted")
except:
    print("no file found")
    pass

#%%
time.sleep(60)
kill_browser("chrome")