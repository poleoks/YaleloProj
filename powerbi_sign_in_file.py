#%% import modules
from credentials import *   # expected to provide d365_user, d365_pass (and possibly other imports)

# Explicit imports so the script doesn't depend on what credentials.py happens to export
import os
import time
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#%% Browser setup
HEADLESS = False   # Keep False for the first run so you can complete any MFA prompt manually

PROFILE_DIR = r"C:\Users\Administrator\selenium_profile"   # dedicated profile (NOT your default Chrome profile)
DOWNLOAD_DIR = r"C:\Users\Administrator\Downloads"        # must match the folder used in download_address

options = webdriver.ChromeOptions()
options.add_argument(f"--user-data-dir={PROFILE_DIR}")
options.add_argument("--profile-directory=Default")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_experimental_option("detach", True)   # keep Chrome open after the script ends
options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
})

if HEADLESS:
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")

# If CHROME_PATH is the path to chrome.exe (NOT a profile/flag), uncomment these two lines:
# from chrome_configuration import CHROME_PATH
# options.binary_location = CHROME_PATH

# Selenium 4.6+ downloads/locates the matching ChromeDriver automatically (Selenium Manager)
browser = webdriver.Chrome(options=options)

#%% Sign in to D365
browser.get('https://fw-d365-prod.operations.dynamics.com/')
if not HEADLESS:
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
            EC.presence_of_element_located((By.XPATH, '//*[@id="otherTileText"]'))
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
            EC.presence_of_element_located((By.XPATH, '//*[@id="i0118"]'))
        )
        # send the user's password
        password.send_keys(d365_pass)
        time.sleep(2)
        # find and click the sign in button
        submit = browser.find_element(By.XPATH, '//*[@id="idSIButton9"]')
        submit.click()
    except:
        try:
            # find the password field
            password = WebDriverWait(browser, 2).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="i0118"]'))
            )
            password.send_keys(d365_pass)
            time.sleep(2)
            submit = browser.find_element(By.XPATH, '//*[@id="idSIButton9"]')
            submit.click()
        except:
            try:
                WebDriverWait(browser, 12).until(
                    EC.presence_of_element_located((By.XPATH, f'//*[@data-test-id="{d365_user}"]'))
                ).click()
                print("Profile saved, not signed in")
                password = WebDriverWait(browser, 2).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="i0118"]'))
                )
                password.send_keys(d365_pass)
                time.sleep(2)
                submit = browser.find_element(By.XPATH, '//*[@id="idSIButton9"]')
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


#%% Power BI helper functions
f_path = []


def pbi_sign_in(repo_url):
    browser.get(repo_url)
    try:
        reset_button = WebDriverWait(browser, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@data-testid="reset-to-default-btn"]'))
        )
        reset_button.click()
        ok_button = WebDriverWait(browser, 15).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="okButton"]'))
        )
        ok_button.click()
        print("reset activated")
    except:
        print("reset not found")

    print("Expanding View")
    try:
        full_screen_btn = WebDriverWait(browser, 60).until(
            EC.element_to_be_clickable((By.ID, "fullScreen"))
        )
        try:
            full_screen_btn.click()
        except Exception:
            # fallback if another element is overlapping the button
            browser.execute_script("arguments[0].click();", full_screen_btn)
        print("Expanded view")
    except Exception:
        print("Full screen button not found, URL:", browser.current_url)
        browser.save_screenshot(r"C:\Users\Administrator\pbi_debug.png")
        raise

    time.sleep(10)


def pbi_export(url, download_address):
    """Export a visual's data from Power BI to Excel, read it into a DataFrame, then delete the file."""
    try:
        os.remove(download_address)
        print(f"{download_address} file removed")
    except:
        print("no file found")
        pass

    browser.get(url)
    hover_element = WebDriverWait(browser, 5 * 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[ @role="presentation" and @class="top-viewport"]'))
    )

    ActionChains(browser).double_click(hover_element).perform()

    WebDriverWait(browser, 5 * 60).until(
        EC.presence_of_element_located((
            By.XPATH,
            '//*[@data-testid="visual-more-options-btn" and @class="vcMenuBtn" '
            'and @aria-expanded="false" and @aria-label="More options"]'
        ))
    ).click()

    # Click Export data
    WebDriverWait(browser, 5 * 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@title="Export data"]'))
    ).click()

    # Click Export icon
    WebDriverWait(browser, 5 * 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@aria-label="Export"]'))
    ).click()

    # Wait for the file to download
    WebDriverWait(browser, 5 * 60).until(
        lambda driver: len(glob.glob(f"{download_address}")) > len(f_path)
    )
    time.sleep(2)

    xx = pd.read_excel(f"{download_address}").iloc[:-1, :]
    print(xx.shape)
    print(xx.head())
    try:
        os.remove(download_address)
        print("file download completed!!!")
        print(f"{download_address} removed!")
    except:
        print("file not found")
        pass

    return xx