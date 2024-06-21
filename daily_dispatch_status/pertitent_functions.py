### whatsapp connection Modules
import time
import os
from selenium import webdriver
##
# from selenium import webdriver
from selenium.webdriver.chrome.service import Service
##

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from tableau_api_lib import TableauServerConnection
import pandas as pd
import pyperclip
from selenium.webdriver.common.keys import Keys
from datetime import datetime, timedelta
from config import CHROME_PROFILE_PATH
current_time = str(datetime.now().time())
#%%
def whatsapp(groups_path,saved_image_path):
    # options = Options()
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument(CHROME_PROFILE_PATH)
    options.add_argument('--no-sandbox')
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--disable-extensions")

    # Disable webdriver flags or you will be easily detectable
    options.add_argument("--disable-blink-features")
    options.add_argument("--disable-blink-features=AutomationControlled")
    browser = webdriver.Chrome(service=service, options=options)

    # Old Chromedriver settings
    # browser = webdriver.Chrome(
        # executable_path=r"C:/Users/Data.Team/Desktop/datateamug/chromedriver",options = options)
        # executable_path=r"C:/Users/Data-Team-UG/Desktop/datateamug/Chromedriver/chromedriver",options = options)
        # executable_path=r"C:/Users/Data-Team-UG/Downloads/chromedriver-win64 (1).zip/chromedriver-win64/chromedriver",options =options)

    # browser.maximize_window()
    
    browser.get('https://web.whatsapp.com/')
    with open(groups_path,'r', encoding='utf8') as f:
        groups = [group.strip() for group in f.readlines()]



    #%%
    # with open('/Users/joseph.bwete/Myspace/Research_folder/Whats_selenium/msg.txt','r', encoding='utf8') as f:
    #     msg = f.read()
        
    # image_path = '/Users/joseph.bwete/Desktop/Daily_Sales_Dash.png'
    image_path = saved_image_path
    
    # image_path = '/Users/joseph.bwete/Downloads/Daily Sales updates_new.png'
    #%%
    # time.sleep(20)
    time.sleep(5)

    completed_groups = []

    for group in groups:
        # image_path = '{group}'.format(group)
        search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
        search_box = WebDriverWait(browser, 500).until(
            EC.presence_of_element_located((By.XPATH, search_xpath))
        )
        search_box.clear()
        time.sleep(1)
        pyperclip.copy(group)
        search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
        time.sleep(2)
        group_xpath = f'//span[@title="{group}"]'
        group_title = browser.find_element(by=By.XPATH, value=group_xpath)
        # group_title = browser.find_elements(by=By.XPATH, value=group_xpath)
        # Test 3
        browser.execute_script("arguments[0].scrollIntoView();", group_title)

        # Alternative incase of failure  
        # group_title = browser.find_element(by=By.XPATH, value=group_xpath)
        # browser.execute_script("arguments[0].click();", group_title)   
           
        group_title.click()
        time.sleep(5)

        #Incase of Text
        # input_xpath = '//div[@contenteditable="true"][@data-tab="10"]'
        # input_box = browser.find_element(by=By.XPATH, value=input_xpath)

        # pyperclip.copy(msg)
        # input_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
        # input_box.send_keys(Keys.ENTER)
        # time.sleep(1)

        attachment_box = browser.find_element(by=By.XPATH, value='//div[@title="Attach"]')
        attachment_box.click()
        time.sleep(3)

        image_box = browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_box.send_keys(image_path)
        time.sleep(3)

        #new [Attach message to image]
        # txt_xpath = '//div[@contenteditable="true"][@data-tab="10"]'
        # txt_xpath = '//div[@contenteditable="true"][@data-testid="media-caption-input-container"]'
        txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
        txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

        # Change here [change commentary]
        if  current_time < '02:00':
            print(current_time)
            pyperclip.copy("Hello *{0}*, Here is the midnight update. Let's close any pending sales # *Maja Manji*".format(group))
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            time.sleep(5)
            completed_groups.append(group)

        elif current_time < '04:00':
            print(current_time)
            pyperclip.copy("Hello *{0}*, Well done and thank you for the great push. Here is the last update of sales yesterday. Let's prepare to push harder today, good night! # *Maja Manji*".format(group))
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            time.sleep(5)
            completed_groups.append(group)
        elif  current_time < '10:55':
            print(current_time)
            pyperclip.copy("Hello *{0}*, These are the sales update as it closed yesterday, let's give it our best today # *Maja Manji*".format(group))
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            time.sleep(5)
            completed_groups.append(group)
        
        elif current_time < '13:00':
            # pyperclip.copy("Hello *{0}*, Apologies the sales update went out when the tables hadn't fully updated.. Happy Hunting. # *Maja Manji*".format(group))
            pyperclip.copy("Hello *{0}*, Are we all in field? These are the results so far. Happy Hunting. # *Maja Manji*".format(group))
            #pyperclip.copy("Hi *{0}*, Which POS is trying to match us in today's sales??? C'mon team let's show them what we've got!!! # *Cheetah Sprint Time*".format(group))
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            time.sleep(5)
            completed_groups.append(group)

        elif current_time < '15:00':
            pyperclip.copy("Hello *{0}*, These are the daily results so far. Lets Keep pushing and hit our Targets. # *Maja Manji*".format(group))
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            time.sleep(5)
            completed_groups.append(group)
        elif current_time < '18:00':
            pyperclip.copy("Hello *{0}*, Please see the 5:00 PM update of the daily sales. Lets push harder. # *Maja Manji*".format(group))
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            # C:/Users/Data-Team-UG/Desktop/datateamug/sales/daily_sales_test_file.py
            # C:/Users/Data-Team-UG/Desktop/datateamug/sales/daily_sales_test_file.py
            time.sleep(5)
            completed_groups.append(group)
        elif current_time < '20:00':
            pyperclip.copy("Hello *{0}*, Here are the results so far, unleash your inner cheetah and sprint - let's be quick and unstoppable!!!! # *Cheetah Speed*".format(group))
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            time.sleep(5)
            completed_groups.append(group)
        elif current_time < '23:59':
            print(current_time)
            pyperclip.copy("Hello *{0}*, Any pending sales out there? Here are the results so far. Let's close all pending sales # *Maja Manji*".format(group))
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            time.sleep(15)
            completed_groups.append(group)

    return completed_groups


#%%
def sending_error(utc_time,Fenix_DB_duration,Sales_Cases,Sales_Cases_h,Sales_Order_Items,Sales_Order_Items_h):
    # options = Options()
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument(CHROME_PROFILE_PATH)
    options.add_argument('--no-sandbox')
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--disable-extensions")

    # Disable webdriver flags or you will be easily detectable
    options.add_argument("--disable-blink-features")
    options.add_argument("--disable-blink-features=AutomationControlled")
    browser = webdriver.Chrome(service=service, options=options)
    
    browser.get('https://web.whatsapp.com/')

    search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
    search_box = WebDriverWait(browser, 500).until(
        EC.presence_of_element_located((By.XPATH, search_xpath))
    )
    search_box.clear()
    time.sleep(1)
    group = "Business Strategy Team - EEA"
    pyperclip.copy(group)
    search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    time.sleep(2)
    group_xpath = f'//span[@title="{group}"]'
    group_title = browser.find_element(by=By.XPATH, value=group_xpath) 
    browser.execute_script("arguments[0].scrollIntoView();", group_title)
    group_title.click()
    time.sleep(2)

    msg = f"""Hi Team,as at *{utc_time}* utc time./n/nThe following Tables are not updated:/n*FenixDB*: {Fenix_DB_duration} ;/n*Sales cases* ({Sales_Cases},{Sales_Cases_h}) ;/n*Sales Order items* ({Sales_Order_Items},{Sales_Order_Items_h})"""
    #Incase of Text
    # @data-tab="3"
    input_xpath = '//div[@contenteditable="true"][@title="Type a message"]'
    input_box = browser.find_element(by=By.XPATH, value=input_xpath)
    pyperclip.copy(msg)
    input_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    input_box.send_keys(Keys.ENTER)
    time.sleep(5)

#%%

def size_error(image_size):
    # options = Options()
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument(CHROME_PROFILE_PATH)
    options.add_argument('--no-sandbox')
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--disable-extensions")

    # Disable webdriver flags or you will be easily detectable
    options.add_argument("--disable-blink-features")
    options.add_argument("--disable-blink-features=AutomationControlled")
    browser = webdriver.Chrome(service=service, options=options)
    
    browser.get('https://web.whatsapp.com/')

    search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
    search_box = WebDriverWait(browser, 500).until(
        EC.presence_of_element_located((By.XPATH, search_xpath))
    )
    search_box.clear()
    time.sleep(1)
    group = "Business Strategy Team - EEA"
    pyperclip.copy(group)
    search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    time.sleep(2)
    group_xpath = f'//span[@title="{group}"]'
    group_title = browser.find_element(by=By.XPATH, value=group_xpath) 
    browser.execute_script("arguments[0].scrollIntoView();", group_title)
    group_title.click()
    time.sleep(2)

    msg = f"""Hi Team, image downloaded is too small or not in compatible format,{image_size})"""
    #Incase of Text
    # @data-tab="3"
    input_xpath = '//div[@contenteditable="true"][@title="Type a message"]'
    input_box = browser.find_element(by=By.XPATH, value=input_xpath)
    pyperclip.copy(msg)
    input_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    input_box.send_keys(Keys.ENTER)
    time.sleep(5)