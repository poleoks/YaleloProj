import pyperclip
import sys
import glob
from Feed_Inventory_Movement import *
sys.path.append("C:/Users/Administrator/Documents/Python_Automations/")
from whatsapp_configuration import CHROME_PROFILE_PATH_Yvone

#INSTANTIATE WHATSAPP
#%%
Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)
Options.add_argument(CHROME_PROFILE_PATH_Yvone)

chrome_install = ChromeDriverManager().install()

folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

Service = webdriver.ChromeService(chromedriver_path)

browser=webdriver.Chrome(service=Service, options = Options)

growout_pic = glob.glob('C:/Users/Administrator/Documents/Python_Automations/Production/G*.png')#:
# screenshots.append((ss_path, page_info['message']))
growout_pics =' '.join(growout_pic)
# growout_pics = []
# growout_pics.append(growout_pic)


# import glob

nursery_pic = glob.glob('C:/Users/Administrator/Documents/Python_Automations/Production/N*.png') + glob.glob('C:/Users/Administrator/Documents/Python_Automations/Production/B*.png')

nursery_pics =' '.join(nursery_pic)
# nursery_pics = [] 
# nursery_pics.append(nursery_pic)
print(nursery_pics)
time.sleep(1)

# ss_path = os.path.join(current_dir, "/", ss_name)
# print(f"This is the picture path: {ss_path}")
# browser=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
# browser.delete_all_cookies()
#%%
groups_path='C:/Users/Administrator/Documents/Python_Automations/Production/feed_grps.txt'

browser.get('https://web.whatsapp.com/')
with open(groups_path,'r', encoding='utf8') as f:
    groups = [group.strip() for group in f.readlines()]

# yesterday_date= yesterday_date.strptime('%d/%m/%y')
time.sleep(5)

for group in groups:
    # ss_path = '{group}'.format(group)
    search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
    search_box = WebDriverWait(browser, 500).until(
        EC.presence_of_element_located((By.XPATH, search_xpath))
    )
    search_box.clear()
    time.sleep(3)
    pyperclip.copy(group)
    search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    time.sleep(2)
    
    group_xpath = f'//span[@title="{group}"]'
    group_title = browser.find_element(by=By.XPATH, value=group_xpath)
    # group_title = browser.find_elements(by=By.XPATH, value=group_xpath)
    # Test 3
    
    browser.execute_script("arguments[0].scrollIntoView();", group_title)
    group_title.click()
    time.sleep(5)
    if group =='Grow Out feeding':
        for  z in growout_pic:
            # zz = z[0][-10:-4]
            zz = os.path.splitext(os.path.basename(z))[0]
            attachment_box = browser.find_element(by=By.XPATH, value='//button[@title="Attach"]')
            attachment_box.click()
            time.sleep(3)

            image_box =browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
            
            image_box.send_keys(z)
            
            txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
            txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

            pyperclip.copy(f'Mortality Status GrowOut Platforms')
            # pyperclip.copy(f'Platform: {yy}')
            txt_box.send_keys(f'Platform: {zz}')
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            time.sleep(10)

    else:
        for y in nursery_pic:
            # yy = y[0][-10:-4]
            yy= os.path.splitext(os.path.basename(y))[0]
            attachment_box = browser.find_element(by=By.XPATH, value='//button[@title="Attach"]')
            attachment_box.click()
            time.sleep(3)

            image_box =browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
            image_box.send_keys(y)

            txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
            txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

            pyperclip.copy(f'Mortality and Feed Status Nursery Platforms')
            # pyperclip.copy(f'Platform: {yy}')
            txt_box.send_keys(f'Platform: {yy}')
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
            # txt_box.send_keys(Keys.ENTER)
            #stop
            send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
            send_btn.click()
            time.sleep(10)

#%%
print("image sent via whatsapp....browser exiting!!!")
time.sleep(3)
browser.quit()

time.sleep(5)
#%%
import os
import glob
for h in glob.glob('C:/Users/Administrator/Documents/Python_Automations/Production/*Feed.png'):
    os.remove(h)
# %%
