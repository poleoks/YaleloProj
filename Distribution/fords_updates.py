#%%
import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/')
from powerbi_sign_in_file import *
import pyperclip
from datetime import date,datetime, timedelta

my_date = date.today()
#get email lists
#%%
#
#navigate the report and take screenshot
current_dir= "C:/Users/Administrator/Documents/Python_Automations/Distribution/"
ss_name ='distribution.png'
ss_path=current_dir+ss_name
# ss_path = os.path.join(current_dir, "/", ss_name)
# print(f"This is the picture path: {ss_path}")

fords_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/c334f3016733ad04622c"

download_address=glob.glob("C:/Users/Administrator/Downloads/data"+"*xlsx")

file_path=[]
    # load url
def export(url):

    for h in download_address:
        os.remove(h)

    browser.get(url)
    hover_element=WebDriverWait(browser, 5*60).until(
        EC.presence_of_element_located((By.XPATH,'//*[ @role="presentation" and @class="top-viewport"]'))
    )

    ActionChains(browser).double_click(hover_element).perform()

    WebDriverWait(browser, 60).until(
            EC.presence_of_element_located((By.XPATH,'//*[@data-testid="visual-more-options-btn" and @class="vcMenuBtn" and @aria-expanded="false" and @aria-label="More options"]'))
        ).click()

    #Click Export data
    WebDriverWait(browser, 3*60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@title="Export data"]'))
    ).click()

    #Click Export Icon
    WebDriverWait(browser, 3*60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@aria-label="Export"]'))
    ).click()

    # Wait for the file to download (this might need adjustments based on download time)
    WebDriverWait(browser, 3*60).until(
        lambda driver: len(glob.glob("C:/Users/Administrator/Downloads/data*.xlsx")) > len(file_path)
    )
    return print("file download completed!!!")

#%%
# read files
export(fords_url)

#%%
tr_in = pd.read_excel("C:/Users/Administrator/Downloads/data.xlsx")
tr_in.to_excel('5_weeks_harvest.xlsx', sheet_name="5_weeks", index=False)
browser.quit()
# log-in to email server
"""
####################################################################
# Email With Attachments Python Script
# 
####################################################################
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
email_list = ['fasega@yalelo.ug','pokuttu@yalelo.ug']
# 
# Define the password (better to reference externally)
pswd = pbi_pass # As shown in the video this password is now dead, left in as example only
# 
print("Connecting to server...")
TIE_server = smtplib.SMTP(smtp_server, smtp_port)
TIE_server.starttls()
TIE_server.login(email_from, pbi_pass_email)
print("Succesfully connected to server")
#%%
subject = "Harvest Report [5 Weeks]"
# 
# Prepare the body of the email
body = f"""
Hello Ford,

Please find harvest report for the last 5 weeks attached.

Regards,
Pole
"""
# Make a MIME object to define parts of the email
msg = MIMEMultipart()
msg['From'] = email_from
msg['Subject'] = subject
msg['To'] = ', '.join(email_list)

msg.attach(MIMEText(body, 'plain'))
# 
# Define the file to attach

filename = "5_weeks_harvest.xlsx"
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
print(f"email sent!")
#%%
print("All Emails Sent successfully")
#clear space
TIE_server.quit()
#%%
for k in glob.glob("P:/Pertinent Files/" + "*.xlsx"):
    os.remove(k)#
# 
print("All files removed from repository")
# %%#
# 