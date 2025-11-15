#%%
from distribution_updates_backend import *
import pyperclip
# 
#%%
#sending whatsapp
# #navigate the report and take screenshot
current_dir= "C:/Users/Administrator/Documents/Python_Automations/Distribution"

ss_name ='distribution.png'
ss_path=current_dir+ss_name

from whatsapp_configuration import CHROME_PROFILE_PATH
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
image_path = ss_path
#%%

# def msg():
#     if day_of_the_week_num==0:
#         return "MTD dispatch report"    
#     elif day_of_the_week_num==6:
#         return "This week's dispatch report"
#     else:
#         return ""
    
# print(msg())

    #%%
time.sleep(5)

for group in groups:
    # image_path = '{group}'.format(group)
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

    attachment_box = browser.find_element(by=By.XPATH, value='//div[@title="Attach"]')
    attachment_box.click()
    time.sleep(3)

    image_box =browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_box.send_keys(image_path)
    time.sleep(3)


#%%
    txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
    txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

    # pyperclip.copy("Good Morning *{0}*, Thank you for closing the week last night - here is the report of dispatch (D365). Let's do better this week. *SPEED IS IN OUR DNA*".format(group))
    
    pyperclip.copy('msg()')
    txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    # txt_box.send_keys(Keys.ENTER)
    #stop
    send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    send_btn.click()
    time.sleep(5)

#%%

print("image sent via whatsapp....driver exiting!!!")

time.sleep(3)
browser.quit()