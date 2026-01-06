##%import module
from credentials import  *

#%%
Options=webdriver.ChromeOptions()
Options.headless = True  # Enable headless mode
Options.add_experimental_option("detach", True)
#INSTANTIATE WHATSAPP
# sys.path.insert(0,'C:/Users/Administrator/Documents/Python_Automations/Distribution/')
from chrome_configuration import CHROME_PATH
Options.add_argument(CHROME_PATH)
chrome_install = ChromeDriverManager().install()

folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

Service = webdriver.ChromeService(chromedriver_path)
browser=webdriver.Chrome(service=Service, options = Options)

#sign-in to power bi
browser.get('https://fw-d365-prod.operations.dynamics.com/')
browser.maximize_window()

try:
    WebDriverWait(browser, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//*[@class="navigation-bar-userInitials"]'))
    )
    print("It was already signed in")
    time.sleep(3)
except:
    print("Wasn't already signed in")
    try:
        WebDriverWait(browser, 5).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="otherTileText"]'))
        ).click()
    except:
        pass
    try:
        time.sleep(3)
        email = browser.find_element(By.XPATH, '//*[@id="i0116"]')
        email.click()
        print("start from scratch")
        time.sleep(2)
        email.send_keys(d365_user)
        time.sleep(2)
        email.send_keys(Keys.ENTER)
        time.sleep(2)
        password = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="i0118"]'))
        )
        # then we send our user's password 
        password.send_keys(d365_pass)
        # after we find sign in button above
        time.sleep(2)
        submit = browser.find_element(By.XPATH,'//*[@id="idSIButton9"]')
        # then we click to submit button
        submit.click()
    except:
        try:
            #  now we need to find password field
            password = WebDriverWait(browser, 2).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="i0118"]'))
            )
            # then we send our user's password 
            password.send_keys(d365_pass)
            # after we find sign in button above
            time.sleep(2)
            submit = browser.find_element(By.XPATH,'//*[@id="idSIButton9"]')
            # then we click to submit button
            submit.click()
        except:
            try:
                WebDriverWait(browser, 12).until(
                    EC.presence_of_element_located((By.XPATH,f'//*[@data-test-id="{d365_user}"]'))
                ).click()
                print("Profile saved, not signed in")
                # time.sleep(9)
                # now we need to find password field
                password = WebDriverWait(browser, 2).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="i0118"]'))
                )
                # then we send our user's password 
                password.send_keys(d365_pass)
                # after we find sign in button above
                time.sleep(2)
                submit = browser.find_element(By.XPATH,'//*[@id="idSIButton9"]')
                # then we click to submit button
                submit.click()
            except:
                pass
    try:
        print("click no")
        WebDriverWait(browser, 5).until(
            EC.presence_of_all_elements_located((By.XPATH, '//*[@value="No"]'))
        )[0].click()
    except:
        pass

       
#%%
#PowerBI Report Sign in Function
f_path = []
def pbi_sign_in(repo_url):
    browser.get(repo_url)
    try:
          reset_button= WebDriverWait(browser, 60*5).until(
                EC.presence_of_element_located((By.XPATH,'//*[@data-testid="reset-to-default-btn"]'))
                )
          
          reset_button.click()
          ok_button = WebDriverWait(browser, 15).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="okButton"]'))
                )
          ok_button.click()
          print("reset activated")
    except:
          print("reset not found")
          pass
    
    print("Expanding View")
    # Click the expand button
    view_button= WebDriverWait(browser, 15).until(
          EC.presence_of_element_located((By.XPATH,'//*[@data-testid="app-bar-view-menu-btn"]'))
    )
    view_button.send_keys(Keys.CONTROL + Keys.SHIFT + 'f')
    
    print("Expanded view")

    time.sleep(10)
def pbi_export(url,download_address):
    try:
        os.remove(download_address)
        print(f"{download_address} file removed")
    except:
        print("no file found")
        pass

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
        lambda driver: len(glob.glob(f"{download_address}")) > len(f_path)
    )
    time.sleep(2)

    xx= pd.read_excel(f"{download_address}").iloc[:-1,:]
    print(xx.shape)
    os.remove(download_address)
    print("file download completed!!!")
    print(f"{download_address} removed!")
    
    return xx
