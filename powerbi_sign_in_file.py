##%import module
from credentials import  *

sys.path.append('E:/Python_Automations/')
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
        time.sleep(3)
        email = browser.find_element(By.XPATH, '//*[@id="i0116"]')
        email.click()
        print("start from scratch")
        time.sleep(1)
        email.send_keys(d365_user)
        email.send_keys(Keys.ENTER)
        time.sleep(1)
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
                WebDriverWait(browser, 2).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@data-test-id="d365@yalelo.ke"]'))
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

    time.sleep(10)