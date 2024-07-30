import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from credentials import *
from datetime import datetime, timedelta
import calendar
import pyperclip

#%%
# GET DATES TO BE USED IN REPORT FILTERING

my_date=datetime.today() 
yesterday= datetime.today() - timedelta(days=1)
print(yesterday)

day_of_the_week=calendar.day_name[my_date.weekday()]
day_of_the_week_num=datetime.today().weekday() #0 for Monday, 1 for Tuesday
month_start_date =my_date.replace(day=1).strftime('%m/%d/%Y')
month_end_date = my_date.strftime('%m/%d/%Y')
yesterday_date = yesterday.strftime('%m/%d/%Y')


print(yesterday_date)
week_start_date=my_date - timedelta(days=my_date.weekday())
week_end_date=week_start_date+timedelta(days=6)
sd_date=week_start_date-timedelta(days=7)
ed_date=week_end_date- timedelta(days=7)

sd_date=sd_date.strftime('%m/%d/%Y')
ed_date=ed_date.strftime('%m/%d/%Y')

week_start_date=week_start_date.strftime('%m/%d/%Y')
week_end_date=week_end_date.strftime('%m/%d/%Y')
#%%
# INSTANTIATE CHROME
Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)
driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
#%%
# powerbi login url
driver.get('https://app.powerbi.com/?noSignUpCheck=1')

#Expand window
driver.maximize_window()
# wait for page load
time.sleep(5)

# find email area using xpath
email=driver.find_element(By.XPATH,"//input[@class='form-control ltr_override input ext-input text-box ext-text-box']")
email.send_keys(pbi_user)

# proceed to password input by clicking button
submit = driver.find_element("id",'idSIButton9')
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
no = driver.find_element("id",'idBtn_Back')
# then we click to no
no.click()

time.sleep(5)

#%%
#navigate specific report by getting its url
report_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/327e7fb0-8c3c-4596-bdd3-98e41603c2a6/ReportSection2e1d7001a05a5136ed5d"
driver.get(report_url)

# wait and locate report load the END DATE
ed=WebDriverWait(driver, 5*60).until(
        EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"End date") and @aria-description="Enter date in M/d/yyyy format"]'))
    )
#clear the search box
ed.clear()
#paste date
ed.send_keys(yesterday_date)
# press enter button
ed.send_keys(Keys.ENTER)

#click enter date on the calendar pop-up grid
driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()

time.sleep(3)

#Wait and locate the report to load the Start Date
sd=WebDriverWait(driver, 5*60).until(
    EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"Start date") and @aria-description="Enter date in M/d/yyyy format"]'))
)
#clear search box
sd.clear()
#paste the date
sd.send_keys(month_start_date)

#press Enter Key
sd.send_keys(Keys.ENTER)
time.sleep(3)

#click on the entered date in the calendar pop-up grid
driver.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()
#Expand view option
driver.find_element(By.XPATH,"//button[contains(@class, 'mat-menu-trigger') and contains(@aria-label, 'View')]").click()

# select full screen
expand_button = driver.find_element(By.XPATH, "//button[contains(@class, 'appBarMatMenu') and contains(@title, 'Open in full-screen mode')]")

# Click the expand button
expand_button.click()

time.sleep(5)

#check if filter exists and hide, else ignore
try:
    driver.find_element(By.XPATH,'//*[@aria-label="Show/hide filter pane" and @aria-expanded="true"]').click()
    print("button pressed to hide filter pane")
except:
    print('Filter already hidden')
time.sleep(3)



# # Take a screenshot of the report and save it to a file
ss_path="P:/Pertinent Files/Python/scripts/daily_dispatch_status/" +'distribution.png'

# driver.save_screenshot(ss_path)
# print("successfully taken screenshot!!!")

# time.sleep(5)
# driver.quit()


#%%
#sending whatsapp
#configure path to save login instance for whatsapp
from config import CHROME_PROFILE_PATH

#INSTANTIATE WHATSAPP
Options.add_argument(CHROME_PROFILE_PATH)
Options.add_experimental_option("detach", True)

#create chrome engine
browser =webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)

#maximize
browser.maximize_window()

#get names of whatsapp contacts/groups from text file
groups_path='P:/Pertinent Files/Python/scripts/daily_dispatch_status/groups_all.txt'
browser.get('https://web.whatsapp.com/')
with open(groups_path,'r', encoding='utf8') as f:
    groups = [group.strip() for group in f.readlines()]
image_path = ss_path

    #%%
time.sleep(10)

# loop whatsapp for groups extracted from the groups_all.txt
for group in groups:
    
    #search the contact search bar
    search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
    search_box = WebDriverWait(browser, 500).until(
        EC.presence_of_element_located((By.XPATH, search_xpath))
    )

    #clear box
    search_box.clear()
    time.sleep(3)

    #copy group name
    pyperclip.copy(group)
    
    #paste the group name
    search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
    time.sleep(2)

    # scroll to the name returned and click
    group_xpath = f'//span[@title="{group}"]'
    group_title = browser.find_element(by=By.XPATH, value=group_xpath)
    browser.execute_script("arguments[0].scrollIntoView();", group_title)
    group_title.click()
    time.sleep(5)

    #click add attachment
    attachment_box = browser.find_element(by=By.XPATH, value='//div[@title="Attach"]')
    attachment_box.click()
    time.sleep(3)
    
    #select add photos/videos
    image_box =browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_box.send_keys(image_path)
    time.sleep(3)

    # search the text path
    txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
    txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

    #copy the message from the msg() function
    pyperclip.copy('Hello Team, Here is the MTD latest standing on stockloss and all related movements\n👉🏽 *Ensure your "Received Quantity, SalesQty, Closing Balance" match this report.*\nIn-box in case of any concerns.')

    #paste msg
    txt_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users

    # click send
    send_btn = browser.find_element(by=By.XPATH, value='//span[@data-icon="send"]')
    send_btn.click()
    time.sleep(5)

#%%

print("image sent via whatsapp....browser exiting!!!")

time.sleep(3)
browser.quit()