from credentials import *
import os
import datetime
import time
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
from dateutil.relativedelta import relativedelta

Options=Options()
Options.add_experimental_option("detach", True)

#%%
#get all start and end dates for each months
def get_first_and_last_days_last_12_months():
        # Create an empty DataFrame to store the appended data
        df = pd.DataFrame()
        browser=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)
        browser.delete_all_cookies()
        browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=LedgerTrialBalanceListPage')
        browser.maximize_window()

        br=WebDriverWait(browser,60).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="i0116"]'))
        )
        br.send_keys(pbi_user)
        br.send_keys(Keys.ENTER)

        br=WebDriverWait(browser,60).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="i0118"]'))
        )
        br.clear()                        
        br.send_keys(pbi_pass)
        br.send_keys(Keys.ENTER)

        WebDriverWait(browser, 60*2).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, 'popupShadow'))
        )

        br=WebDriverWait(browser,60).until(
            EC.presence_of_element_located((By.XPATH,'//*[@value="Sign in"]'))
            )
        br.click()
        br=WebDriverWait(browser,60).until(
            EC.presence_of_element_located((By.XPATH,'//*[@value="No"]'))
            )
        br.click()
        print('sign-in completed, wait starts')
#%% 
        for i in range(12):
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
                      
            print(f"Next loop for the {j(i)} Month, has started")
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

            # //*[contains(@id,"ledgertrialbalancelistpage") and contains(@id,"EndDate_input")]
            br=WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[contains(@id,"ledgertrialbalancelistpage") and contains(@id,"EndDate_input")]'))
                )
            br.clear()
            time.sleep(1)
            br.send_keys(last_day)
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


            br=WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[contains(@id,"FormControlFieldSelector_Selected") and contains(@id,"container")]/div/span'))
            )
            br.click()
            print('rows filtered')

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

            WebDriverWait(browser,60).until(
                 EC.presence_of_element_located((By.XPATH,'//*[@class="messageBar-message" and contains(text(),"We finished your export")]'))
            )

#%%
            # Use glob to find files matching the pattern "Trial balance*" in the Downloads folder
            downloads_path = "C:/Users/Pole Okuttu/Downloads/"

            # Use glob to find files matching the pattern "Trial balance*" in the Downloads folder
            for file in glob.glob(os.path.join(downloads_path, "Trial balance*")):
                try:
                    data = pd.read_excel(file)
                    print(f"{j(i)} month data shape is {data.shape}")
                    df = pd.concat([df, data])
                    print(f"consolidated data shape is {df.shape}")
                    print(first_day, last_day)
                    os.remove(file)
                    print('File deleted:', file)
                except Exception as e:
                    print("Error:", e)
            time.sleep(3)
        print("Compilation completed")
        browser.quit()
        df.to_excel('Consolidated_tb.xlsx', index=False)

get_first_and_last_days_last_12_months()

#%%

"""
######################################################################
# Email With Attachments Python Script

######################################################################
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Setup port number and server name

smtp_port = 587                 # Standard secure SMTP port
smtp_server = "smtp-mail.outlook.com"  # Google SMTP Server

# Set up the email lists
email_from = pbi_user
email_list = ["pokuttu@yalelo.ug"]
# email_list = ["pokuttu@yalelo.ug","knyeko@yalelo.ug","rnabukeera@yalelo.ug","aoriide@yalelo.ug","alakica@yalelo.ug"]

# Define the password (better to reference externally)
pswd = pbi_pass # As shown in the video this password is now dead, left in as example only


# name the email subject
subject = f"Trial Balance Extract"

import sys 
import os
new_path='P:/Pertinent Files/Python/scripts/'
# sys.path.insert(0, new_path)
os.chdir(new_path)
# Define the email function (dont call it email!)
def send_emails(email_list):

    # Prepare the body of the email
    body = """
    Hello Team,
    
    Please find an updated record of Trial Balance for the past 12 months attached.

    Regards,
    Audit Team
    """

    # Make a MIME object to define parts of the email
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['Subject'] = subject
    # recipients = '; '.join(email_list)
     # Add CC recipients
    # cc_recipients = email_list
    # msg['CC'] = ', '.join(email_list)
    msg['To'] = ', '.join(email_list)
    # Attach the body of the message
    msg.attach(MIMEText(body, 'plain'))

    # Define the file to attach
    filename = "Consolidated_tb.xlsx"

    # Open the file in python as a binary
    with open(filename, 'rb') as attachment:
        # Encode as base 64
        attachment_package = MIMEBase('application', 'octet-stream')
        attachment_package.set_payload(attachment.read())
        encoders.encode_base64(attachment_package)
        attachment_package.add_header('Content-Disposition', f"attachment; filename= {filename}")
        msg.attach(attachment_package)

    # Connect with the server
    print("Connecting to server...")
    TIE_server = smtplib.SMTP(smtp_server, smtp_port)
    TIE_server.starttls()
    TIE_server.login(email_from, pswd)
    print("Succesfully connected to server")
    print()

    # Send emails to all recipients at once
    print("Sending emails...")
    TIE_server.sendmail(email_from,email_list, msg.as_string())
    print("All emails sent")
    print()

    # Close the port
    TIE_server.quit()

# Run the function
send_emails(email_list)

# %%
