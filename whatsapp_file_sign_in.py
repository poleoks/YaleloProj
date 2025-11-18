import sys
# dir_path = "D:/YU/ScriptCodez/"
dir_path = "C:/Users/Administrator/Documents/Python_Automations/"
sys.path.append(dir_path)
from credentials import  *
from whatsapp_configuration import *



def whatsapp_sign_in(whatsapp_owner):
    def kill_browser(process_name="chrome"):

        """Kill all browser and driver processes."""
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if process_name.lower() in proc.info['name'].lower():
                    os.kill(proc.info['pid'], signal.SIGTERM)
            except Exception:
                pass

    try:
        kill_browser("chrome")
        print("Running browser killed")
    except:
        print("No Running browser, proceed")
        pass
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
   
    return browser



def whatsapp_share(recipient_group, wsp_message,wsp_file, directory, profile_owner):
    #whatsapp
    browser=whatsapp_sign_in(profile_owner)
    group_data = {}
    for group, message, file in zip(recipient_group, wsp_message, wsp_file):
        if group not in group_data:
            # Initialize the group with a list of (message, file) tuples
            group_data[group] = []
        # Append the message and file for this group
        group_data[group].append((message, file))
    for group_name, messages_and_files in group_data.items():
        # --- 1. SEARCH & SELECT GROUP (Done once per group) ---
        print(f"Searching for group: {group_name}")
        search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
        search_box = WebDriverWait(browser, 500).until(
            EC.presence_of_element_located((By.XPATH, search_xpath))
        )
        search_box.clear()
        time.sleep(3)

        # Use group_name for search/paste
        pyperclip.copy(group_name)
        search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
        time.sleep(2)

        group_xpath = f'//*[@title="{group_name}"]'
        group_title = WebDriverWait(browser, 500).until(
            EC.presence_of_element_located((By.XPATH, group_xpath))
        )
        
        browser.execute_script("arguments[0].scrollIntoView();", group_title)
        group_title.click()
        time.sleep(5)

        # --- 2. INNER LOOP: SEND ALL FILES AND MESSAGES ---
        # messages_and_files is a list of (message, file) tuples for the current group
        for wsp_message, wsp_file in messages_and_files:
            print(f"Sending file '{wsp_file}' and message to {group_name}...")

            # --- A. Attach File ---
            attachment_box = browser.find_element(by=By.XPATH, value="//span[@data-icon='plus-rounded']")
            attachment_box.click()
            time.sleep(3)

            image_box = browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
            # Assuming 'directory' is defined outside this scope
            full_path = directory + wsp_file
            image_box.send_keys(full_path) 
            time.sleep(3)

            # --- B. Add Caption/Message and Send ---
            txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
            txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

            # Copy and paste the specific message for this file
            pyperclip.copy(f'{wsp_message}')
            txt_box.send_keys(Keys.SHIFT, Keys.INSERT)
            
            # Send message (which also sends the attached file)
            txt_box.send_keys(Keys.ENTER)
            time.sleep(5)
            
        print(f"✅ Finished sending all items to {group_name}.")

    print("Image sent via WhatsApp. Browser exiting (or proceeding with next steps)...")
