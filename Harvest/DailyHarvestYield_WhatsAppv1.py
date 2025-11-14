#%%##
#remove file if existent
import os
from harvest_data_logmanager_posts import *
from datetime import datetime, timedelta
import dataframe_image as dfi
import pyperclip
#%%sending whatsapp
import sys
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import matplotlib.pyplot as plt

## navigate the report and take screenshot
sys.path.append("C:/Users/Administrator/Documents/Python_Automations/")
from whatsapp_config_pole import CHROME_PROFILE_PATH
from gmail_sender import *
from credentials import *
#%%
try:
    os.remove('Harvest/harvest.png')
    print('file deleted')
    time.sleep(2)
except:
    print('file DNE')

#%%
# print(df.tail())
try:
    today = datetime.datetime.today()# - timedelta(days=3)
    last_time = df.tail(1)['datetime'].min()
    print(today,last_time)
    currentdatetime = (datetime.datetime.today() - timedelta(days=19)).strftime('%Y-%m-%d %H:%M:%S')
    currentdatetime = today  #.strftime('%Y-%m-%d %H:%M:%S')
    currentdate = today.strftime('%Y-%m-%d')
    df['datetime'] = pd.to_datetime(df['datetime'])
    df['netweight'] = df['netweight'].str.replace('kg','').astype('float')
    df['batch_number'] = df['batch_number'].str[:4] + "("+ df['batch_number'].str[7:] +")"
    df['number_of_pieces'] = df['number_of_pieces'].astype('int')
    df['date'] = df['datetime'].dt.strftime('%Y-%m-%d')
    df['timedifference_mins'] = (abs(df['datetime'] - currentdatetime).dt.total_seconds() / 60).astype('int')
    last_date =  df.tail(1)['date'].min()
    # print(currentdate, last_date)
    (xrows,ycols) = df.shape
    print(df.shape)

    #%%
    # Define the custom order for materials
    material_order = {
        'SXXL': 1,
        'XXL': 2,
        'Extra Large': 3,
        'Large': 4,
        'Medium': 5,
        'Small': 6,
        'Extra Small': 7,
        'HSXXL': 8,
        'HXXL': 9,
        'HExtra Large': 10,
        'HLarge': 11,
        'HMedium': 12,
        'HSmall': 13,
        'Below 150': 14,
        'Mortality': 15,
        'Sick Fish': 16,
        'Culled': 17
    }

    batch_number_order = {
        batch: idx for idx, batch in enumerate(sorted(df['batch_number'].unique()), start=1)
    }

        
    if (df.sort_values(by='datetime').tail(1)['timedifference_mins'].min() < 120):# and xrows > 10:
        # Group by and aggregate
        df2 = df[df['date'] == currentdate][['batch_number', 'material', 'netweight', 'number_of_pieces']].groupby(['batch_number', 'material']).agg(
            total_weight_kg=('netweight', 'sum'),
            total_pieces=('number_of_pieces', 'sum'),
            total_crates=('batch_number', 'count')
        )

        s = df2.groupby(level=0).sum()
        s.index = pd.MultiIndex.from_product([s.index, ['**BATCH SUMMARY**'], ['']])

        # Concatenate and apply custom sorting
        dd = pd.concat([df2, s])

        # Custom sorting based on both batch_number and material
        dd = dd.sort_index(level=[0, 1], key=lambda idx: (
            idx.map(batch_number_order) if idx.name == 'batch_number' else idx.map(material_order)
        ))

        # Add 'Grand Total' row
        dd.loc['Grand Total', :] = df2.sum().values

        # Additional formatting
        dd['ABW (g)'] = (dd.total_weight_kg * 1000 / dd.total_pieces).astype('int').apply(lambda x: '{:,}'.format(x))
        dd['total_crates'] = dd['total_crates'].astype('int').apply(lambda x: '{:,}'.format(x))
        dd['total_pieces'] = dd['total_pieces'].astype('int').apply(lambda x: '{:,}'.format(x))
        dd['total_weight_kg'] = dd['total_weight_kg'].astype('float').round(2).apply(lambda x: '{:,}'.format(x))

        print(dd)

        # Reset index and adjust appearance
        dd = dd.reset_index()
        time.sleep(5)

        # Replace duplicate 'batch_number' values with empty strings
        dd['batch_number'] = dd['batch_number'].mask(dd['batch_number'].duplicated(), '')
        dd.rename(columns={'level_1': 'SKU'}, inplace=True)

        # Style and export
        df_styled = dd.style.hide(axis='index')

        # Export the DataFrame as an image
        dfi.export(df_styled, 'C:/Users/Administrator/Documents/Python_Automations/Harvest/harvest.png', table_conversion='matplotlib', max_rows=None, max_cols=None, fontsize=12)
        time.sleep(5)


        time.sleep(5)
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
        groups_path='C:/Users/Administrator/Documents/Python_Automations/Distribution/click.txt'
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
                time.sleep(2)

                # Optional code for attachment
                pyperclip.copy(f'Latest harvest report as at (Last Crate Weighed): {last_time}')
                time.sleep(2)
    
                attachment_box = browser.find_element(by=By.XPATH, value="//span[@data-icon='plus-rounded']")
                attachment_box.click()
                # time.sleep(3)
                time.sleep(2)
                image_box = browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')

                image_box.send_keys('C:/Users/Administrator/Documents/Python_Automations/Harvest/harvest.png')                
                
                time.sleep(2)

                txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
                txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

                txt_box.send_keys(Keys.SHIFT, Keys.INSERT)  # Use Keys.CONTROL + "v" for Windows users

                txt_box.send_keys(Keys.ENTER)
                
                print("Message sent via WhatsApp... Driver exiting!!!")
                time.sleep(300)
            except Exception as e:
                print(f"An error occurred: {e}")
                time.sleep(3)
        
        browser.quit()
    else:
        print("No Latest Harvest Data!")
        pass
        # email_function(pbi_user,pbi_pass_email,['pokuttu@yalelo.ug'],'Harvest Update On Whatsapp Has Failed',
        #        'Hello, the harvest subscription for Whatsapp has failed to go. Please look into it.'
        #        )
except:
    email_function(pbi_user,'Harvest Update On Whatsapp Has Failed',
               'Hello, the harvest subscription for Whatsapp has failed to go. Please look into it.',None
               )