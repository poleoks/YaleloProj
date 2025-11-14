#%%
from harvest_data_logmanager1 import *
from datetime import datetime, timedelta
import dataframe_image as dfi
import pyperclip
#sending whatsapp
import sys
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# #navigate the report and take screenshot

sys.path.append("C:/Users/Administrator/Documents/Python_Automations/")
from whatsapp_configuration import CHROME_PROFILE_PATH

#%%
print(df.tail())
# currentdatetime = (datetime.today() - timedelta(days=19)).strftime('%Y-%m-%d %H:%M:%S')
currentdatetime = datetime.today()#.strftime('%Y-%m-%d %H:%M:%S')
currentdate = datetime.today().strftime('%Y-%m-%d')
df['datetime'] = pd.to_datetime(df['datetime'])
df['netweight'] = df['netweight'].astype('float')
df['number_of_pieces'] = df['number_of_pieces'].astype('int')
df['date'] = df['datetime'].dt.strftime('%Y-%m-%d')
df['timedifference_mins'] = ((currentdatetime - df['datetime']).dt.total_seconds() / 60).astype('int')
# %%
# Group by and aggregate
df2 = df[df['date'] == currentdate][['batch_number', 'material', 'netweight', 'number_of_pieces']].groupby(['batch_number', 'material']).agg(
    total_weight=('netweight', 'sum'),
    total_pieces=('number_of_pieces', 'sum'),
    total_crates=('batch_number', 'count')
)
df2['ABW'] = (df2.total_weight * 1000/ df2.total_pieces).astype('int')
df3 = df[df['date'] == currentdate][['batch_number', 'material', 'netweight', 'number_of_pieces']].groupby(['batch_number']).agg(
    total_weight=('netweight', 'sum'),
    total_pieces=('number_of_pieces', 'sum'),
    total_crates=('batch_number', 'count')
)

df3['ABW'] = (df3.total_weight * 1000/ df3.total_pieces).astype('int')
df4 =df3.sum()
df4['ABW'] = (df4.total_weight * 1000/ df4.total_pieces).astype('int')
df4 = pd.DataFrame(df4)
df4.reset_index()
df4 = df4.T
print(df4)
print(df3)
print(df2)

df2.dfi.export('split_df_harvest.png')
df3.dfi.export('per_batch_df_harvest.png')
df4.dfi.export('overall_df_harvest.png')

current_dir= "C:/Users/Administrator/Documents/Python_Automations"
for i in glob.glob(f'{current_dir}/*df_harvest.png'):
    print(i)
#%%
#INSTANTIATE WHATSAPP
Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)
Options.add_argument(CHROME_PROFILE_PATH)
chrome_install = ChromeDriverManager().install()

folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

Service = webdriver.ChromeService(chromedriver_path)
browser=webdriver.Chrome(service=Service, options=Options)

# browser =webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
browser.delete_all_cookies()
browser.maximize_window()
#%%

groups_path='C:/Users/Administrator/Documents/Python_Automations/Distribution/click_pole.txt'
browser.get('https://web.whatsapp.com/')
with open(groups_path,'r', encoding='utf8') as f:
    groups = [group.strip() for group in f.readlines()]

# Make sure df2 is defined before using it
for group in groups:
    try:
        search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
        search_box = WebDriverWait(browser, 500).until(
            EC.presence_of_element_located((By.XPATH, search_xpath))
        )
        search_box.clear()
        time.sleep(3)
        pyperclip.copy(group)
        search_box.send_keys(Keys.SHIFT, Keys.INSERT)  # Use Keys.CONTROL + "v" for Windows users
        time.sleep(2)
        
        group_xpath = f'//span[@title="{group}"]'
        group_title = browser.find_element(by=By.XPATH, value=group_xpath)
        browser.execute_script("arguments[0].scrollIntoView();", group_title)
        group_title.click()
        time.sleep(5)
        l=0
        for i in glob.glob(f'{current_dir}/*df_harvest.png'):
            def msg(i):
                if l==0:
                    return "overall summary"
                elif l==1:
                    return "batch summary"
                else:
                    return "general split"
                    # Optional code for attachment
            pyperclip.copy(msg(i))
            attachment_box = browser.find_element(by=By.XPATH, value='//div[@title="Attach"]')
            attachment_box.click()
            time.sleep(2)
            image_box = browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
            image_box.send_keys(i)
            time.sleep(2)

            txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
            txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

            
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT)  # Use Keys.CONTROL + "v" for Windows users
            txt_box.send_keys(Keys.ENTER)
            # send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            # send_btn.click()
            l+=1
            time.sleep(5)

    except Exception as e:
        print(f"An error occurred: {e}")

print("Message sent via WhatsApp... Driver exiting!!!")

time.sleep(3)
browser.quit()