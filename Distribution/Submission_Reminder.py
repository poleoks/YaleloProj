#%%
import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/')
from credentials import  *
from powerbi_sign_in_file import *
import glob
import time
import pandas as pd 
# from PIL import Image
import pyperclip

#%%

from datetime import date,datetime, timedelta
import calendar
my_date = date.today()
#get email lists


report_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/327e7fb0-8c3c-4596-bdd3-98e41603c2a6/67fa000c96323c50965b?ctid=683df1d2-484b-4bb5-8ca6-1cdbf2b98d28"
download_address=glob.glob("C:/Users/Administrator/Downloads/data"+"*xlsx")

file_path=[]

def export(url):

    for h in download_address:
        os.remove(h)

    browser.get(url)
    hover_element=WebDriverWait(browser, 5*60).until(
        EC.presence_of_element_located((By.XPATH,'//*[ @role="presentation" and @class="top-viewport"]'))
    )

    ActionChains(browser).double_click(hover_element).perform()

    WebDriverWait(browser, 5*60).until(
            EC.presence_of_element_located((By.XPATH,'//*[@data-testid="visual-more-options-btn" and @class="vcMenuBtn" and @aria-expanded="false" and @aria-label="More options"]'))
        ).click()

    #Click Export data
    WebDriverWait(browser, 5*60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@title="Export data"]'))
    ).click()

    #Click Export Icon
    WebDriverWait(browser, 5*60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@aria-label="Export"]'))
    ).click()

    # Wait for the file to download (this might need adjustments based on download time)
    WebDriverWait(browser, 5*60).until(
        lambda driver: len(glob.glob("C:/Users/Administrator/Downloads/data*.xlsx")) > len(file_path)
    )
    time.sleep(2)
    
    browser.quit()

export(report_url)
#%%
df = pd.read_excel("C:/Users/Administrator/Downloads/data.xlsx")
df = df.iloc[:-1,:]
df['Closing Balance'] = df['Closing Balance'].fillna('Blank')
df = df[df['Closing Balance']=='Blank']
pending_submitters = df['Warehouse'].dropna().tolist()
print(len(pending_submitters)<=0)

