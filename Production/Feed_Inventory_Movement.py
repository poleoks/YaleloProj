# #%%
from selenium.webdriver.common.keys import Keys
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
current_dir= "C:/Users/Administrator/Documents/Python_Automations/Production/"

# ss_name ='Feed.png'
ss_path=current_dir#+ss_name

report_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/e33a4d76-14e2-475d-b035-62fb4a47e444/65f141bc08c07d1050a3"

browser.get(report_url)
# wait for report load

#%%
##wait for a max of 5 mins until full load

## select the last week calendar performance from the date filter
# ed=WebDriverWait(browser, 5*60).until(
#         EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"End date") and @aria-description="Enter date in M/d/yyyy format"]'))
#     )
# ed.clear()
# ed.send_keys(yesterday_date)
# ed.send_keys(Keys.ENTER)

# browser.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()

# time.sleep(3)

# sd=WebDriverWait(browser, 5*60).until(
#     EC.presence_of_element_located((By.XPATH,'//*[contains(@aria-label,"Start date") and @aria-description="Enter date in M/d/yyyy format"]'))
# )
# sd.clear()
# sd.send_keys(yesterday_date)
# sd.send_keys(Keys.ENTER)



# browser.find_element(By.XPATH,'//*[@class="date-cell themableBackgroundColorSelected date-selected"]').click()
#Expand view to full screen
WebDriverWait(browser,20).until(
    EC.presence_of_element_located((By.XPATH,"//button[contains(@class, 'mat-menu-trigger') and contains(@aria-label, 'View')]"))
).click()
time.sleep(1)
expand_button = browser.find_element(By.XPATH, "//button[contains(@class, 'appBarMatMenu') and contains(@title, 'Open in full-screen mode')]")

# Click the expand button
expand_button.click()

time.sleep(5)

try:
    browser.find_element(By.XPATH,'//*[@aria-label="Show/hide filter pane" and @aria-expanded="true"]').click()
    print("button pressed to hide filter pane")
except:
    print('Filter already hidden')
time.sleep(5)

for jj in ['BS','GA','GB','GC','GD','GE','NA','NB']:
    
    try:
        browser.find_element(By.XPATH, f"//*[@title='{jj}']").click()
        time.sleep(5)
        # Take a screenshot of the report and save it to a file
        browser.save_screenshot(f"{ss_path}/{jj}_Feed.png")
        print(f"{jj} screenshot taken!")
    except:
        pass

    time.sleep(2)
print("successfully taken screenshots!!!")

time.sleep(5)
browser.quit()