#%%
import pyperclip
import sys
import os
from Submission_Reminder import *
# sys.path.insert(1,'D:/Pertinent Files/Python/scripts/')

active_warehouse = [
    {'WarehouseId' : 'Bulaga', 'Email':'bulagaduuka@yalelo.ug'},
    {'WarehouseId' : 'ButembeSal', 'Email':'butembesales@yalelo.ug'},
    {'WarehouseId' : 'Bunamwaya', 'Email':'bunamwayastore@yalelo.ug'},
    {'WarehouseId' : 'Busia', 'Email':'busiastore@yalelo.ug'},
    {'WarehouseId' : 'Bwaise', 'Email':'bwaisestore@yalelo.ug'},
    {'WarehouseId' : 'Distrbtion', 'Email':'mbiryetega@yalelo.ug; jnanyonjo@yalelo.ug'},
    {'WarehouseId' : 'KyebandoDC', 'Email':'mbiryetega@yalelo.ug; jnanyonjo@yalelo.ug'},
    {'WarehouseId' : 'Jinja', 'Email':'jinjastore@yalelo.ug'},
    {'WarehouseId' : 'Kafunta', 'Email':'kafuntastore@yalelo.ug'},
    {'WarehouseId' : 'Kasangati', 'Email':'kasangatistore@yalelo.ug'},
    {'WarehouseId' : 'Kasubi', 'Email':'kasubistore@yalelo.ug'},
    {'WarehouseId' : 'Kawempe', 'Email':'kawempestore@yalelo.ug'},
    {'WarehouseId' : 'Kibuli', 'Email':'kibulistore@yalelo.ug'},
    {'WarehouseId' : 'Kibuye', 'Email':'kibuyestore@yalelo.ug'},
    {'WarehouseId' : 'Kireka', 'Email':'kirekastore@yalelo.ug'},
    {'WarehouseId' : 'Kyaliwajal', 'Email':'kyaliwajjalastore@yalelo.ug'},
    {'WarehouseId' : 'Kyambogo', 'Email':'kyambogostore@yalelo.ug'},
    {'WarehouseId' : 'Kyengera-R', 'Email':'kyengerastore@yalelo.ug'},
    {'WarehouseId' : 'Malaba', 'Email':'malababorder@yalelo.ug'},
    {'WarehouseId' : 'Mbale', 'Email':'mbalestore@yalelo.ug'},
    {'WarehouseId' : 'Mpondwe', 'Email':'mpondweborder@yalelo.ug'},
    {'WarehouseId' : 'Mukono', 'Email':'mukonoshop@yalelo.ug'},
    {'WarehouseId' : 'Mutungo', 'Email':'mutungostore@yalelo.ug'},
    {'WarehouseId' : 'Nansana', 'Email':'nansanastore@yalelo.ug'},
    {'WarehouseId' : 'Natete', 'Email':'nateetestore@yalelo.ug'},
    {'WarehouseId' : 'Ntinda', 'Email':'ntindastore@yalelo.ug'},
    {'WarehouseId' : 'Nyahuka', 'Email':'nyahukaborder@yalelo.ug'}
    ]
active_warehouse = pd.DataFrame(active_warehouse)
active_warehouse = active_warehouse[active_warehouse['WarehouseId'].isin(pending_submitters)]
print(active_warehouse)


#%%
download_address=glob.glob("C:/Users/Pole Okuttu/Downloads/data"+"*xlsx")
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
pswd = pbi_pass_email # As shown in the video this password is now dead, left in as example only
# 
print("Connecting to server...")
TIE_server = smtplib.SMTP(smtp_server, smtp_port)
TIE_server.starttls()
TIE_server.login(email_from, pswd)
print("Succesfully connected to server")
#%%
for i,e in zip(active_warehouse['WarehouseId'].to_list(), active_warehouse['Email'].to_list()):
    email_list = [f"{e}","pokuttu@yalelo.ug", "alakica@yalelo.ug","bmbaraga@yalelo.ug","anagudi@yalelo.ug","bkabuye@yalelo.ug"]
   # name the email subject
# 
    subject = f"Attention {i}! Notice to Fill In Your Stock Balance Data"

    # Prepare the body of the email
    body = f"""
    Hello {i} Store,

    Please use this link to fill your stock balance:
    https://forms.microsoft.com/r/As4eEheKZj

    This must be done within the next 30 minutes if it is to reflect in the next Stock Movement Updates.

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

    TIE_server.sendmail(email_from,email_list, msg.as_string())
    print(f"{i} email sent!")
#%%
print("All Emails Sent successfully")
#clear space
TIE_server.quit()