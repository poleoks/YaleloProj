from credentials import  *
from whatsapp_configuration import *


def wsp_sgn(recipient_group, wsp_message,wsp_file, whatsapp_owner, directory):
    #whatsapp
    Options=webdriver.ChromeOptions()
    Options.add_experimental_option("detach", True)
    Options.add_argument(whatsapp_owner)
    Options.headless = True  # Enable headless mode

    chrome_install = ChromeDriverManager().install()

    folder = os.path.dirname(chrome_install)
    chromedriver_path = os.path.join(folder, "chromedriver.exe")

    Service = webdriver.ChromeService(chromedriver_path)

    browser=webdriver.Chrome(service=Service, options = Options)
    
    browser.get('https://web.whatsapp.com/')
    search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
    search_box = WebDriverWait(browser, 500).until(
            EC.presence_of_element_located((By.XPATH, search_xpath))
            )
    search_box.clear()
    time.sleep(3)
    pyperclip.copy(recipient_group)
    search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    time.sleep(2)

    group_xpath = f'//*[@title="{recipient_group}"]'
    group_title = WebDriverWait(browser, 500).until(
            EC.presence_of_element_located((By.XPATH, group_xpath))
    )
    # group_title = browser.find_elements(by=By.XPATH, value=group_xpath)
    # Test 3

    browser.execute_script("arguments[0].scrollIntoView();", group_title)
    group_title.click()
    time.sleep(5)

    attachment_box = browser.find_element(by=By.XPATH, value="//span[@data-icon='plus-rounded']")
    attachment_box.click()
    time.sleep(3)

    os.chdir(directory)

    image_box =browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_box.send_keys(wsp_file)
    time.sleep(3)

    #%%
    txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
    txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

    # pyperclip.copy("Hi Team, appologies there was a glitch in the earlier sent report. Here is the updated on")
    pyperclip.copy(f'{wsp_message}')
    txt_box.send_keys(Keys.SHIFT, Keys.INSERT)
    #send message
    txt_box.send_keys(Keys.ENTER)
    time.sleep(5)
    print("image sent via whatsapp....browser exiting!!!")
#%%
groups_t = ["Pole", "Pole Mifi", "Pole"]
messages_t = [
    ['Platform: GA_Feed', 'Platform: GB_Feed', 'Platform: GC_Feed', 'Platform: GD_Feed', 'Platform: GE_Feed', 'Platform: GF_Feed'],
    ['Platform: NA_Feed', 'Platform: NB_Feed', 'Platform: BS_Feed'],
    ['Platform: GA_Feed', 'Platform: GB_Feed', 'Platform: GC_Feed', 'Platform: GD_Feed', 'Platform: GE_Feed', 'Platform: GF_Feed',
     'Platform: NA_Feed', 'Platform: NB_Feed', 'Platform: BS_Feed']
]
files_t = [
        ['GA_Feed.png', 'GB_Feed.png', 'GC_Feed.png', 'GD_Feed.png', 'GE_Feed.png', 'GF_Feed.png'],
        ['NA_Feed.png', 'NB_Feed.png', 'BS_Feed.png'],
        ['GA_Feed.png', 'GB_Feed.png', 'GC_Feed.png', 'GD_Feed.png', 'GE_Feed.png', 'GF_Feed.png',
         'NA_Feed.png', 'NB_Feed.png', 'BS_Feed.png']
]

# Send to each group one by one
file_path = "C:/Users/Administrator/Documents/Python_Automations/Production/"
os.chdir(file_path)
for grp, msg_list, file_list in zip(groups_t, messages_t, files_t):
    print(f"🔹 Sending to group: {grp}")
    
    # Loop through all messages and files for this group
    for msg, file in zip(msg_list, file_list):
        wsp_sgn(grp, msg, file, Pole_dp, file_path)
        print(f"   ✅ Sent {file} with message: {msg}")
    
    print(f"✅ Finished sending to {grp}\n")