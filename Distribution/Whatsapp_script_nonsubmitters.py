#%%
import pyperclip
import sys
import os
from Submission_Reminder import *
sys.path.insert(1,'D:/Pertinent Files/Python/scripts/')

if len(pending_submitters) > 0:
    #%%
    from whatsapp_configuration import CHROME_PROFILE_PATH_R
    #INSTANTIATE WHATSAPP
    #%%
    Options=webdriver.ChromeOptions()
    Options.add_experimental_option("detach", True)
    Options.add_argument(CHROME_PROFILE_PATH_R)

    chrome_install = ChromeDriverManager().install()

    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")

    Service = webdriver.ChromeService(chromedriver_path)

    driver=webdriver.Chrome(service=Service, options=Options)
    # driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
    # driver.delete_all_cookies()
    #%%
    time.sleep(2)
    #
    groups_path='C:/Users/Administrator/Documents/Python_Automations/Distribution/click_madinah.txt'
    driver.get('https://web.whatsapp.com/')

    driver.maximize_window()
    time.sleep(3)
    with open(groups_path,'r', encoding='utf8') as f:
        groups = [group.strip() for group in f.readlines()]
    ss_path="C:/Users/Administrator/Documents/Python_Automations/Distribution/distribution.png"
    #%%
    for group in groups:
        # ss_path = '{group}'.format(group)
        search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
        search_box = WebDriverWait(driver, 500).until(
            EC.presence_of_element_located((By.XPATH, search_xpath))
        )
        search_box.clear()
        time.sleep(2)
        pyperclip.copy(group)
        search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
        time.sleep(2)
        group_xpath = f'//span[@title="{group}"]'
        group_title = driver.find_element(by=By.XPATH, value=group_xpath)
        # group_title = driver.find_elements(by=By.XPATH, value=group_xpath)
        # Test 3
        driver.execute_script("arguments[0].scrollIntoView();", group_title) 
        
        group_title.click()
        time.sleep(2)
    
        txt_xpath = '//*[@aria-placeholder="Type a message"]'
        txt_box = driver.find_element(by=By.XPATH, value=txt_xpath)
        
        msg = '\n'.join(pending_submitters)
        txt_box.send_keys(f"The following stores haven't submitted closing balance \n{msg} \nPlease submit by 6:55 AM for it to be included in the next update")
        #stop
        time.sleep(3)
        
        send_btn = driver.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
        send_btn.click()

        print("All shops have filled the data")

    #%%
    time.sleep(5)
    driver.quit()
else:
    print("All shops have submitted closing balance")