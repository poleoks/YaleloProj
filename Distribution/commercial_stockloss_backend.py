# #%%
import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/')
from powerbi_sign_in_file import *
#%%

from datetime import datetime, timedelta
import calendar

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


print(f"Month start date: {sd_date}")
print(f"Month end date: {ed_date}")

print(f"week start date: {week_start_date}")
print(f"week end date: {week_end_date}")

current_time = str(datetime.now().time())

#%%
#navigate the report and take screenshot
current_dir= "C:/Users/Administrator/Documents/Python_Automations/Distribution/"

ss_name ='commercial_stockloss.png'
ss_path=current_dir+ss_name
# ss_path = os.path.join(current_dir, "/", ss_name)
print(f"This is the picture path: {ss_path}")


report_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/327e7fb0-8c3c-4596-bdd3-98e41603c2a6/ReportSection2e1d7001a05a5136ed5d"


# 
browser.get(report_url)
# wait for report load

#%%
#wait for a max of 5 mins until full load

# select the last week calendar performance from the date filter
ed=WebDriverWait(browser, 5*60).until(
        EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"End date") and @aria-description="Enter date in M/d/yyyy format"]'))
    )
ed.clear()
ed.send_keys(yesterday_date)
ed.send_keys(Keys.ENTER)

browser.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()

time.sleep(3)

sd=WebDriverWait(browser, 5*60).until(
    EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"Start date") and @aria-description="Enter date in M/d/yyyy format"]'))
)
sd.clear()
sd.send_keys(month_start_date)
sd.send_keys(Keys.ENTER)

time.sleep(3)


browser.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()
#Expand view to full screen
browser.find_element(By.XPATH,"//button[contains(@class, 'mat-menu-trigger') and contains(@aria-label, 'View')]").click()
expand_button = browser.find_element(By.XPATH, "//button[contains(@class, 'appBarMatMenu') and contains(@title, 'Open in full-screen mode')]")

# Click the expand button
expand_button.click()

time.sleep(5)

try:
    browser.find_element(By.XPATH,'//*[@aria-label="Show/hide filter pane" and @aria-expanded="true"]').click()
    print("button pressed to hide filter pane")
except:
    print('Filter already hidden')
time.sleep(3)
# Take a screenshot of the report and save it to a file
# time.sleep(5)
browser.save_screenshot(ss_path)
print("successfully taken screenshot!!!")

time.sleep(5)
browser.quit()