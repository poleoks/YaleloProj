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
from selenium.common.exceptions import NoSuchElementException
import os
import pandas as pd 
# from PIL import Image
import pyperclip
#%%

from datetime import date,datetime, timedelta
import calendar
my_date = date.today()


day_of_the_week=calendar.day_name[my_date.weekday()]
day_of_the_week_num=datetime.today().weekday() #0 for Monday, 1 for Tuesday
month_start_date =my_date.replace(day=1).strftime('%m/%d/%Y')
month_end_date = my_date.strftime('%m/%d/%Y')


week_start_date=my_date - timedelta(days=my_date.weekday())
week_end_date=week_start_date+timedelta(days=6)

week_start_date=week_start_date.strftime('%m/%d/%Y')
week_end_date=week_end_date.strftime('%m/%d/%Y')

print(f"Month start date: {month_start_date}")
print(f"Month end date: {month_end_date}")

print(f"week start date: {week_start_date}")
print(f"week end date: {week_end_date}")
#%%
current_time = str(datetime.now().time())

# Options.add_experimental_option("detach", True)

Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)

driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
driver.delete_all_cookies()
#%%
# powerbi login url
driver.get('https://app.powerbi.com/?noSignUpCheck=1')
driver.maximize_window()
# wait for page load
time.sleep(5)
# find email area using xpath

# email = driver.find_element("xpath",'//*[@id="i0116"]')

email=driver.find_element(By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']")
# # send power bi login user
pbi_user='d365@yalelo.ke'
pbi_pass='Kenya@YK'
email.send_keys(pbi_user)
#%% find submit button
submit = driver.find_element("id",'idSIButton9')
# # click for submit button
submit.click()

time.sleep(5)
#%%
# now we need to find password field
password = driver.find_element("id",'i0118')
# then we send our user's password 
password.send_keys(pbi_pass)
# after we find sign in button above
submit = driver.find_element("id",'idSIButton9')
# then we click to submit button
submit.click()
# time.sleep(3)
time.sleep(5)
# to complate the login process we need to click no button from above
# find no button by element
no = driver.find_element("id",'idBtn_Back')
# then we click to no
no.click()
time.sleep(5)
# //*[@id="email"]

#%%
#navigate the report and take screenshot
current_dir= "P:/Pertinent Files/Python/scripts/daily_dispatch_status/"
ss_name ='distribution.png'
ss_path=current_dir+ss_name
# ss_path = os.path.join(current_dir, "/", ss_name)
print(f"This is the picture path: {ss_path}")

# report_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSection2436e75a5249a600b5a0"
report_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSectionb07695af0594eb707bed"
# report_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/ReportSection57732dac3e37637503de"
# Navigate to the report URL
# report_url = "https://app.powerbi.com/groups/e8afba4b-7068-47d1-842b-6ab1891c531e/reports/33097f5c-faae-4c44-b82d-791f91c4d09b?ctid=683df1d2-484b-4bb5-8ca6-1cdbf2b98d28&pbi_source=linkShare&bookmarkGuid=2ed61d9b-52e9-40a1-9891-de65ca8e52f0"
driver.get(report_url)
# wait for report load

#%%
#wait for a max of 5 mins until full load
#filter date
if day_of_the_week_num==0:
    sd=WebDriverWait(driver, 5*60).until(
            EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"Start date") and @aria-description="Enter date in M/d/yyyy format"]'))
        )
    sd.clear()
    sd.send_keys(month_start_date)
    sd.send_keys(Keys.ENTER)
    
    driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()
    # select the last week calendar performance from the date filter
    time.sleep(3)

    ed=WebDriverWait(driver, 5*60).until(
            EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"End date") and @aria-description="Enter date in M/d/yyyy format"]'))
        )
    ed.clear()
    ed.send_keys(month_end_date)
    ed.send_keys(Keys.ENTER)

    driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()

    time.sleep(3)
    #Expand view to full screen
    driver.find_element(By.XPATH,"//button[contains(@class, 'mat-menu-trigger') and contains(@aria-label, 'View')]").click()
    expand_button = driver.find_element(By.XPATH, "//button[contains(@class, 'appBarMatMenu') and contains(@title, 'Open in full-screen mode')]")

    # Click the expand button
    expand_button.click()

    time.sleep(5)

    try:
        driver.find_element(By.XPATH,'//*[@aria-label="Show/hide filter pane" and @aria-expanded="true"]').click()
        print("button pressed to hide filter pane")
    except:
        print('Filter already hidden')
    time.sleep(3)
    # Take a screenshot of the report and save it to a file
    # time.sleep(5)
    driver.save_screenshot(ss_path)
    print("successfully taken screenshot!!!")

    time.sleep(3)
    driver.quit()

else:
    sd=WebDriverWait(driver, 5*60).until(
            EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"Start date") and @aria-description="Enter date in M/d/yyyy format"]'))
        )
    sd.clear()
    sd.send_keys(week_start_date)
    sd.send_keys(Keys.ENTER)
    
    driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()
    # select the last week calendar performance from the date filter
    time.sleep(3)

    ed=WebDriverWait(driver, 5*60).until(
            EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"End date") and @aria-description="Enter date in M/d/yyyy format"]'))
        )
    ed.clear()
    ed.send_keys(week_end_date)
    ed.send_keys(Keys.ENTER)

    driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()
    # select the last week calendar performance from the date filter
    time.sleep(3)
    #Expand view to full screen
    driver.find_element(By.XPATH,"//button[contains(@class, 'mat-menu-trigger') and contains(@aria-label, 'View')]").click()
    expand_button = driver.find_element(By.XPATH, "//button[contains(@class, 'appBarMatMenu') and contains(@title, 'Open in full-screen mode')]")

    # Click the expand button
    expand_button.click()

    try:
        driver.find_element(By.XPATH,'//*[@aria-label="Show/hide filter pane" and @aria-expanded="true"]').click()
        print("button pressed to hide filter pane")
    except:
        print('Filter already hidden')
        
    time.sleep(5)
    # Take a screenshot of the report and save it to a file
    # time.sleep(5)
    driver.save_screenshot(ss_path)
    print("successfully taken screenshot!!!")

    time.sleep(30)
    driver.quit()

#%%
#sending whatsapp
from config import CHROME_PROFILE_PATH
Options.add_argument(CHROME_PROFILE_PATH)
Options.add_experimental_option("detach", True)

browser =webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
browser.maximize_window()
#%%

groups_path='P:/Pertinent Files/Python/scripts/daily_dispatch_status/click.txt'
browser.get('https://web.whatsapp.com/')
with open(groups_path,'r', encoding='utf8') as f:
    groups = [group.strip() for group in f.readlines()]
image_path = ss_path

#%%

def msg():
    if day_of_the_week_num==0:
        return "Hello *Distribution Team*, find here last week's dispatch performance so far. Mondays set the tone for the week ahead. Let us start strong.🐋*EeeKyenyanjaa🔥🔥🔥*"
    elif day_of_the_week_num==1:
        return "Good morning *Distribution*, find report for dispatch last night. Tuesday is a day for breakthroughs and innovation. Let us approach our tasks with a creative mindset and enhance our efficiency.💪🏼*EeeKyenyanjaa*🎺🎺🎺"
    elif day_of_the_week_num==2:
        return "Hello *Distribution Team*, Find results of dispatch last night in copy. Be aware that the faster we operate, the more time we create for unforeseen challenges.*EeeKyenyanjaa*🐠🐠🐠"
    elif day_of_the_week_num==3:
        return "Hi *Team*, see attached performance of yesterday. Thursdays are an opportunity to fine tune our strategies.*EeeKyenyanjaa*🐠🐠🐠"
    elif day_of_the_week_num==4:
        return "Hello *Dispatch Team*, find dispatch report as of yesterday. As we dive into Friday tasks, let us maintain the momentum and demonstrate why we are the epitome of speed and precision.*EeeKyenyanjaa*🎺🎺🎺"
    elif day_of_the_week_num==5:
        return "Good morning *Team*, here is dispatch report last night. Our clients rely on us round the clock, and Saturdays are no exception. Let us show them our commitment.*EeeKyenyanjaa*🐟🐟🐟"
    else:
        return "Happy Sunday *Distribution Team✨*, here is report of dispatch yesterday.  Today we recharge our batteries to ensure we enter the new week with a fresh and invigorated mindset.*EeeKyenyanjaa*🐟🐟🐟"
    
print(msg())

    #%%
time.sleep(20)

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
    
    pyperclip.copy(msg())
    txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    # txt_box.send_keys(Keys.ENTER)
    #stop
    send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    send_btn.click()
    time.sleep(5)

#%%

print("image sent via whatsapp....browser exiting!!!")

time.sleep(3)
browser.quit()