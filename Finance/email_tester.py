#%%
import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/')
from credentials import  *
import glob
import os
import pandas as pd 


#%%

from datetime import date,datetime, timedelta
import calendar
my_date = date.today()
#get email lists
#%%
active_warehouse = [
    {'WarehouseId' : 'Distrbtion', 'Email':'pokuttu@yalelo.ug,mkyosiimire@yalelo.ug,eokello@yalelo.ug,ykomwaka@yalelo.ug,rnabukeera@yalelo.ug'}
    ]


active_warehouse = pd.DataFrame(active_warehouse)

#%%
# log-in to email server
"""
######################################################################
# Email With Attachments Python Script
# 
######################################################################
"""
# 
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
# 
# Setup port number and server name
# 
smtp_port = 587                 # Standard secure SMTP port
smtp_server = "smtp-mail.outlook.com"  # Google SMTP Server
# 
# Set up the email lists
email_from = pbi_user
# 
# Define the password (better to reference externally)
pswd = pbi_pass # As shown in the video this password is now dead, left in as example only
# 
print("Connecting to server...")
TIE_server = smtplib.SMTP(smtp_server, smtp_port)
TIE_server.starttls()
TIE_server.login(email_from, pswd)
print("Succesfully connected to server")
#%%
for i,e in zip(active_warehouse['WarehouseId'].to_list(), active_warehouse['Email'].to_list()):
    
    email_list = [email.strip() for email in e.split(',')] 
    # email_list = [f"{e}","poleoks@gmail.com"]
    subject = f"{i} Test Run"
    body = f"""
        Hello {i} Centre,

        Please find your MTD transfers report. There are two sheets (transfers in and transfers out), ensure to check of both.
        The Total Received - Total Shipped Stock = Net Received (shared daily on whatsapp)

        In case of any issues, please reply to pokuttu@yalelo.ug for followup.

        Regards,
        Pole
        """
    # Make a MIME object to define parts of the email
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['Subject'] = subject
    msg['To'] = ', '.join(email_list)
# 
# 
    # Attach the body of the message
    msg.attach(MIMEText(body, 'plain'))
# 
    # Define the file to attach
    filename = f"{i} MTD Stock.xlsx"
# 
    # Open the file in python as a binary
    with open(filename, 'rb') as attachment:
        # Encode as base 64
        attachment_package = MIMEBase('application', 'octet-stream')
        attachment_package.set_payload(attachment.read())
        encoders.encode_base64(attachment_package)
        attachment_package.add_header('Content-Disposition', f"attachment; filename= {filename}")
        msg.attach(attachment_package)
    # Send emails to all recipients at once
    TIE_server.sendmail(email_from,email_list, msg.as_string())
    print(f"{i} email sent!")
#%%
print("All Emails Sent successfully")
#clear space
TIE_server.quit()
#%%
for k in glob.glob("P:/Pertinent Files/"+ "*MTD Stock*" + "*.xlsx"):
    os.remove(k)

print("All files removed from repository")