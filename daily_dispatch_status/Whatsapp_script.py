#%%
#import modules
import selenium
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import pandas as pd 
from PIL import Image
import pyperclip
#%%

from datetime import datetime, timedelta
from config import CHROME_PROFILE_PATH
current_time = str(datetime.now().time())


Options=Options()
Options.add_experimental_option("detach", True)

driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)

#%%
# powerbi login url
# driver.get('https://app.powerbi.com/?noSignUpCheck=1')

#send whatsapp msg
time.sleep(15)
#
groups_path='P:/Pertinent Files/Python/scripts/groups_all.txt'
driver.get('https://web.whatsapp.com/')
driver.maximize_window()
with open(groups_path,'r', encoding='utf8') as f:
    groups = [group.strip() for group in f.readlines()]


ss_path="P:/Pertinent Files/Python/saved_images/distribution.png"
    #%%
    # with open('/Users/joseph.bwete/Myspace/Research_folder/Whats_selenium/msg.txt','r', encoding='utf8') as f:
    #     msg = f.read()
        
    # image_path = '/Users/joseph.bwete/Desktop/Daily_Sales_Dash.png'
image_path = ss_path
    
    # image_path = '/Users/joseph.bwete/Downloads/Daily Sales updates_new.png'
    #%%
    # time.sleep(20)

for group in groups:
    # image_path = '{group}'.format(group)
    search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
    search_box = WebDriverWait(driver, 500).until(
        EC.presence_of_element_located((By.XPATH, search_xpath))
    )
    search_box.clear()
    time.sleep(1)
    pyperclip.copy(group)
    search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    time.sleep(2)
    group_xpath = f'//span[@title="{group}"]'
    group_title = driver.find_element(by=By.XPATH, value=group_xpath)
    # group_title = browser.find_elements(by=By.XPATH, value=group_xpath)
    # Test 3
    driver.execute_script("arguments[0].scrollIntoView();", group_title)

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

    attachment_box = driver.find_element(by=By.XPATH, value='//div[@title="Attach"]')
    attachment_box.click()
    time.sleep(3)

    image_box =driver.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_box.send_keys(image_path)
    time.sleep(3)

    #new [Attach message to image]
    # txt_xpath = '//div[@contenteditable="true"][@data-tab="10"]'
    # txt_xpath = '//div[@contenteditable="true"][@data-testid="media-caption-input-container"]'
    txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
    txt_box = driver.find_element(by=By.XPATH, value=txt_xpath)

    pyperclip.copy("Hello *{0}*, The Magic reporting is here, Now we can push any reports live on whatsapp # *OneTeam💪*".format(group))
    txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    # txt_box.send_keys(Keys.ENTER)
    #stop
    send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    send_btn.click()



    # # Change here [change commentary]
    # if  current_time < '02:00':
    #     print(current_time)
    #     pyperclip.copy("Hello *{0}*, Here is the midnight update. Let's close any pending sales # *Maja Manji*".format(group))
    #     txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    #     # txt_box.send_keys(Keys.ENTER)
    #     #stop
    #     send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    #     send_btn.click()
    #     time.sleep(5)
    #     #completed_groups.append(group)

    # elif current_time < '04:00':
    #     print(current_time)
    #     pyperclip.copy("Hello *{0}*, Well done and thank you for the great push. Here is the last update of sales yesterday. Let's prepare to push harder today, good night! # *Maja Manji*".format(group))
    #     txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    #     # txt_box.send_keys(Keys.ENTER)
    #     #stop
    #     send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    #     send_btn.click()
    #     time.sleep(5)
    #     #completed_groups.append(group)
    # elif  current_time < '10:55':
    #     print(current_time)
    #     pyperclip.copy("Hello *{0}*, These are the sales update as it closed yesterday, let's give it our best today # *Maja Manji*".format(group))
    #     txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    #     # txt_box.send_keys(Keys.ENTER)
    #     #stop
    #     send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    #     send_btn.click()
    #     time.sleep(5)
    #     #completed_groups.append(group)
    
    # elif current_time < '13:00':
    #     # pyperclip.copy("Hello *{0}*, Apologies the sales update went out when the tables hadn't fully updated.. Happy Hunting. # *Maja Manji*".format(group))
    #     pyperclip.copy("Hello *{0}*, Are we all in field? These are the results so far. Happy Hunting. # *Maja Manji*".format(group))
    #     #pyperclip.copy("Hi *{0}*, Which POS is trying to match us in today's sales??? C'mon team let's show them what we've got!!! # *Cheetah Sprint Time*".format(group))
    #     txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    #     # txt_box.send_keys(Keys.ENTER)
    #     #stop
    #     send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    #     send_btn.click()
    #     time.sleep(5)
    #     #completed_groups.append(group)

    # elif current_time < '15:00':
    #     pyperclip.copy("Hello *{0}*, These are the daily results so far. Lets Keep pushing and hit our Targets. # *Maja Manji*".format(group))
    #     txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    #     # txt_box.send_keys(Keys.ENTER)
    #     #stop
    #     send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    #     send_btn.click()
    #     time.sleep(5)
    #     #completed_groups.append(group)
    # elif current_time < '18:00':
    #     pyperclip.copy("Hello *{0}*, Please see the 5:00 PM update of the daily sales. Lets push harder. # *Maja Manji*".format(group))
    #     txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    #     # txt_box.send_keys(Keys.ENTER)
    #     #stop
    #     send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    #     send_btn.click()
    #     # C:/Users/Data-Team-UG/Desktop/datateamug/sales/daily_sales_test_file.py
    #     # C:/Users/Data-Team-UG/Desktop/datateamug/sales/daily_sales_test_file.py
    #     time.sleep(5)
    #     #completed_groups.append(group)
    # elif current_time < '20:00':
    #     pyperclip.copy("Hello *{0}*, Here are the results so far, unleash your inner cheetah and sprint - let's be quick and unstoppable!!!! # *Cheetah Speed*".format(group))
    #     txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    #     # txt_box.send_keys(Keys.ENTER)
    #     #stop
    #     send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    #     send_btn.click()
    #     time.sleep(5)
    #     #completed_groups.append(group)
    # elif current_time < '23:59':
    #     print(current_time)
    #     pyperclip.copy("Hello *{0}*, Any pending sales out there? Here are the results so far. Let's close all pending sales # *Maja Manji*".format(group))
    #     txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    #     # txt_box.send_keys(Keys.ENTER)
    #     #stop
    #     send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    #     send_btn.click()
    #     time.sleep(15)
#
#%%
time.sleep(15)
print("image cropped....browser exiting!!!")

time.sleep(3)
driver.quit()