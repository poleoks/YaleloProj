#%%
all_dispatches = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/2332ae67e68efa624a90"
stock_movement = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/327e7fb0-8c3c-4596-bdd3-98e41603c2a6/ReportSection2e1d7001a05a5136ed5d"
daily_harvest = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/245814b31fad33544821"
import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/')

from credentials import  *
from warehouse_email import *
from email_sender import *
from powerbi_sign_in_file import *
import glob
import time
import pandas as pd 
download_address=glob.glob("C:/Users/Administrator/Downloads/data"+"*xlsx")
#%%


file_path=[]
    # load url
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
    
    return print("file download completed!!!")
#%%
export(all_dispatches)
dispatch = pd.read_excel("C:/Users/Administrator/Downloads/data.xlsx").iloc[:-1,:]
# dispatch.sort_values(by="ReceivingSite", inplace=True)
# dispatch.rename(columns={"ReceivingSite":"DispatchWarehouse"}, inplace=True)
print(dispatch.head(2))
time.sleep(3)

export(daily_harvest)
harvest = pd.read_excel("C:/Users/Administrator/Downloads/data.xlsx").iloc[:-1,:]
print(harvest.head(3))
time.sleep(3)

export(stock_movement)
stock_summary = pd.read_excel("C:/Users/Administrator/Downloads/data.xlsx").iloc[:-1,:]
# stock_summary.sort_values(by="Warehouse", inplace=True)
print(stock_summary.head(4))

browser.quit()
time.sleep(2)

sys.path.append("C:/Users/Administrator/Documents/Python_Automations/Finance/")

with pd.ExcelWriter('stock_flow.xlsx') as wr:
    dispatch.to_excel(wr, sheet_name="dispatch")
    harvest.to_excel(wr, sheet_name="harvest")
    stock_summary.to_excel(wr, sheet_name= "stock_summary")

# email_function(sender_email,sender_password,receipient_email,subject_line,email_body,attachment_path)
email_function(pbi_user,pbi_pass_email,['pokuttu@yalelo.ug'],'stock_flow',
               'This is an ETL for all stock origin and destination',
               'stock_flow.xlsx'
               )