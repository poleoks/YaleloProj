import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/')
from credentials import  *
import os
import datetime
import time
import glob
import pandas as pd

from powerbi_sign_in_file import *
#%%
#get all start and end dates for each months
def get_first_and_last_days_last_12_months():
        # Create an empty DataFrame to store the appended data
        df = pd.DataFrame()
        # browser=webdriver.Chrome(service=Service)
        browser.delete_all_cookies()
        browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=LedgerTrialBalanceListPage')
        browser.maximize_window()
        for i in range(12):
            
            # Use glob to find files matching the pattern "Trial balance*" in the Downloads folder
            downloads_path = "c:/Users/Administrator/Downloads"

            # Use glob to find files matching the pattern "Trial balance*" in the Downloads folder
            for file in glob.glob(os.path.join(downloads_path, "Trial balance*")):
                try:
                    data = pd.read_excel(file)
                    print(f"{j(i)} month data shape is {data.shape}")
                    df = pd.concat([df, data])
                    # print(f"consolidated data shape is {df.shape}")
                    # print(first_day, last_day)
                    os.remove(file)
                except Exception as e:
                    print("Error:", e)
                    
                    
                    
            file_path =[]
            if i>0:
                browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=LedgerTrialBalanceListPage')
            else:
                pass
            j=i+1
            def j(i):
                 if i==0:
                      return '1st'
                 elif i==1:
                      return '2nd'
                 elif i==2:
                      return '3rd'
                 elif i==3:
                      return '4th'
                 elif i==4:
                      return '5th'
                 elif i==5:
                      return '6th'
                 elif i==6:
                      return '7th'
                 elif i==7:
                      return '8th'
                 elif i==8:
                      return '9th'
                 elif i==9:
                      return '10th'
                 elif i==10:
                      return '11th'
                 elif i==11:
                      return '12th'
                 else:
                      pass
                      
            # print(f"Next loop for the {j(i)} Month, has started")
            today = datetime.date.today()
            first_day = today.replace(day=1) - relativedelta(months=i)

            # Calculate the last day of the month
            last_day = (first_day + relativedelta(months=1)) - datetime.timedelta(days=1)
            first_day=first_day.strftime('%m/%d/%Y')
            last_day=last_day.strftime('%m/%d/%Y')

            br=WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[contains(@id,"ledgertrialbalancelistpage") and contains(@id,"StartDate_input")]'))
                )
            br.clear()
            time.sleep(1)
            br.send_keys(first_day)
            br.send_keys(Keys.ENTER)

            # //*[contains(@id,"ledgertrialbalancelistpage") and contains(@id,"EndDate_input")]
            br=WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[contains(@id,"ledgertrialbalancelistpage") and contains(@id,"EndDate_input")]'))
                )
            br.clear()
            time.sleep(1)
            br.send_keys(last_day)
            br.send_keys(Keys.ENTER)
            time.sleep(1)

            br=WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@class="button-label" and contains(text(),"Calculate balances")]'))
                )
            br.click()

            # Wait for the overlay to disappear
            WebDriverWait(browser, 60*2).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, 'popupShadow'))
            )

            WebDriverWait(browser,60).until(
                 EC.presence_of_element_located((By.XPATH,'//*[@id="DimensionAttributeValue01_3_0_0_input"]'))
            )


            # Locate the element
            element = WebDriverWait(browser, 60*2).until(
                EC.presence_of_element_located((By.XPATH, '//*[contains(@id,"ListPageGrid") and @aria-label="Grid options" and contains(@id,"optionsCell_menuButtonButton")]'))
            )

            # Scroll the element
            browser.execute_script("arguments[0].scrollIntoView(true);", element)

            # Wait until the element is clickable
            WebDriverWait(browser, 60*2).until(
                EC.element_to_be_clickable((By.XPATH, '//*[contains(@id,"ListPageGrid") and @aria-label="Grid options" and contains(@id,"optionsCell_menuButtonButton")]'))
            )

            # Click using JavaScript
            browser.execute_script("arguments[0].click();", element)
            # Wait for the overlay to disappear (if any)
            WebDriverWait(browser, 60*2).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, 'popupShadow'))
            )

            # Locate the element
            element = WebDriverWait(browser, 60*2).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="Insert columns..."]'))
            )

            # Scroll the element into view
            browser.execute_script("arguments[0].scrollIntoView(true);", element)

            # Wait until the element is clickable
            WebDriverWait(browser, 60*2).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@aria-label="Insert columns..."]'))
            )

            # Click using JavaScript
            browser.execute_script("arguments[0].click();", element)

            WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,"//span[@class='pivot-label' and contains(text(), 'Recommended fields')]"))
            ).click()
            
            try:
                WebDriverWait(browser,60).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@aria-checked="true" and @aria-label="Select"]'))
                )
            except:
                br=WebDriverWait(browser,60).until(
                    EC.presence_of_element_located((By.XPATH,'//*[contains(@id,"FormControlFieldSelector_Selected") and contains(@id,"container")]/div/span'))
                )
                br.click()

            br=WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[contains(@id,"FormControlFieldSelector") and contains(@id,"OK") and @class="button-label"]'))
            )
            br.click()
            try:
                # Wait until the element is clickable
                element = WebDriverWait(browser, 60*2).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@class='button flyoutButton-Button dynamicsButton' and @name='SystemDefinedOfficeButton' and @aria-label='Open in Microsoft Office']"))
                )
                
                # Scroll the element into view
                browser.execute_script("arguments[0].scrollIntoView(true);", element)
                
                # Adding a small delay to ensure the element is in view
                time.sleep(1)
                
                try:
                    # Attempt to click the element
                    element.click()
                except ElementClickInterceptedException:
                    # If the click is still intercepted, use JavaScript to click
                    browser.execute_script("arguments[0].click();", element)
            except Exception as e:
                # Log the exception or handle it as necessary
                print(f"An error occurred: {e}")

            br=WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="ledgertrialbalancelistpage_1_SystemDefinedOfficeButton_Grid_ListPageGrid_label"]'))
                )
            
            br.click()

            br=WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@data-dyn-controlname="DownloadButton"]'))
            )
            br.click()
            # Wait for the file to download (this might need adjustments based on download time)
            WebDriverWait(browser, 5*60).until(
                lambda driver: len(glob.glob("C:/Users/Administrator/Downloads/Trial balance*.xlsx")) > len(file_path)
            )
            time.sleep(3)
        print("Compilation completed")
        browser.quit()
        df.to_excel('Trial_Balance_1_Year_Extract.xlsx', index=False)

get_first_and_last_days_last_12_months()

#%%
# Set up the email lists
body = "Hello Team, \n\nPlease find an updated record of Trial Balance for the past 12 months attached.\n\nRegards,\nAudit Team"

email_list = "pokuttu@yalelo.ug, alakica@yalelo.ug"

# name the email subject
subject = f"Trial Balance Extract"
new_path='C:/Users/Administrator/Documents/Python_Automations/'
# sys.path.insert(0, new_path)
os.chdir(new_path)

filename = "Trial_Balance_1_Year_Extract.xlsx"


# %%
time.sleep(5)
# Sending email using gmail sender function
from gmail_sender import *
gmail_function(email_list,subject,body,new_path+filename)
# remove the file after sending email
if os.path.exists(new_path+filename):
    os.remove(new_path+filename)
    print(f"Removed the file: {filename} after sending email.")
    
##%
kill_browser("chrome")