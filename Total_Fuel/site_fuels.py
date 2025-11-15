#%%
import glob
import os
file_path = "C:/Users/Administrator/Downloads/fuel_transactions.csv"
attachment_name = "site_fuel_dispensation.xlsx"
#%%
#remove files
try:
    os.remove(file_path)
    print(f"{file_path},removed")
except:
    pass
try:
    for x in glob.glob("C:/Users/Administrator/Documents/Python_Automations/Total_Fuel/*.xlsx"):
        os.remove(x)
        print(f"{x},removed")
except:
    pass
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
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import os
import pandas as pd 
import sys
from datetime import datetime, timedelta

sys.path.append("C:/Users/Administrator/Documents/Python_Automations/")
from credentials import *

Options=Options()
Options.add_experimental_option("detach", True)
driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=Options)

#%%
driver.get('http://7.7.0.2:8081/phpmyadmin/index.php?route=/sql&pos=0&db=fuel&table=fuel_transactions')
time.sleep(2)
driver.maximize_window()

WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'table/export') and contains(@href, 'fuel_transactions')]"))
    ).click()

time.sleep(5)
#%%
dropdown = Select(driver.find_element(By.ID, "plugins"))

# Select by visible text
dropdown.select_by_visible_text("CSV for MS Excel")


#%%

button = driver.find_element(By.XPATH,'//*[@id="buttonGo"]')
# Scroll into view
driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
time.sleep(3)
button.click()
print('clicked')

time.sleep(10)
driver.quit()


# Read DataFrame
df = pd.read_csv(file_path)
df.to_excel(attachment_name,sheet_name="dispensed_fuel")
# Show top rows
print(df.head())
# #%%
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
pswd = pbi_pass_email # As shown in the video this password is now dead, left in as example only


# name the email subject
subject = f"Site Fuel"

def send_emails(email_list):

    # Prepare the body of the email
    body = """
    Hello Team,
    
    Please find an updated record of fuel dispensed from onsite tank.

    Regards,
    Audit Team
    """

    # Make a MIME object to define parts of the email
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['Subject'] = subject
    msg['To'] = ', '.join(email_list)
    # Attach the body of the message
    msg.attach(MIMEText(body, 'plain'))

    # Define the file to attach

    # Open the file in python as a binary
    with open(attachment_name, 'rb') as attachment:
        # Encode as base 64
        attachment_package = MIMEBase('application', 'octet-stream')
        attachment_package.set_payload(attachment.read())
        encoders.encode_base64(attachment_package)
        attachment_package.add_header('Content-Disposition', f"attachment; filename= {attachment_name}")
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


#%%
#Remove file
try:
    os.remove(file_path)
    print(f"{file_path},removed")
except:
    pass

try:
    os.remove(attachment_name)
    print(f"{attachment_name},removed")
except:
    pass
# %%