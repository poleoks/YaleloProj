#%%
import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/')
from credentials import  *
from powerbi_sign_in_file import *
import glob
import time
import pandas as pd 
# from PIL import Image
import pyperclip

#%%

from datetime import date,datetime, timedelta
import calendar
my_date = date.today()

#get email lists

#%%
active_warehouse = [
    # {'WarehouseId' : 'KyebandoDC', 'Email':'pokuttu@yalelo.ug'}
    {'WarehouseId' : 'Bulaga', 'Email':'bulagaduuka@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Bunamwaya', 'Email':'bunamwayastore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Busia', 'Email':'busiastore@yalelo.ug, anagudi@yalelo.ug, wekisa@yalelo.ug, dabigaba@yalelo.ug, ldemol@yalelo.ug,  nbitsinze@yalelo.ug, mnanyibuka@yalelo.ug, pomara@yalelo.ug, pmusoke@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug, rwalusimbi@yalelo.ug, rwalusimbi@yalelo.ug'},
    {'WarehouseId' : 'Bwaise', 'Email':'bwaisestore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'ButembeSal', 'Email':'butembesales@yalelo.ug, anagudi@yalelo.ug, wekisa@yalelo.ug, dabigaba@yalelo.ug, ldemol@yalelo.ug,  nbitsinze@yalelo.ug, mnanyibuka@yalelo.ug, pomara@yalelo.ug, pmusoke@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug, rwalusimbi@yalelo.ug'},
    {'WarehouseId' : 'Jinja V3', 'Email':'jinjastore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kafunta', 'Email':'kafuntastore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kasangati', 'Email':'kasangatistore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kasubi', 'Email':'kasubistore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kawempe V2', 'Email':'kawempestore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kibuli', 'Email':'kibulistore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kibuye', 'Email':'kibuyestore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kireka', 'Email':'kirekastore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kyaliwajal', 'Email':'kyaliwajjalastore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kyambogo', 'Email':'kyambogostore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Kyengera-R', 'Email':'kyengerastore@yalelo.ug, pokuttu@yalelo.ug,  rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Distrbtion', 'Email':'kyengerastore@yalelo.ug, pokuttu@yalelo.ug,  rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Malaba', 'Email':'malababorder@yalelo.ug, anagudi@yalelo.ug, wekisa@yalelo.ug, dabigaba@yalelo.ug, ldemol@yalelo.ug,  nbitsinze@yalelo.ug, mnanyibuka@yalelo.ug, pomara@yalelo.ug, pmusoke@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug, rwalusimbi@yalelo.ug'},
    {'WarehouseId' : 'Mpondwe', 'Email':'mpondweborder@yalelo.ug, anagudi@yalelo.ug, wekisa@yalelo.ug, dabigaba@yalelo.ug, ldemol@yalelo.ug, nbitsinze@yalelo.ug, mnanyibuka@yalelo.ug, pomara@yalelo.ug, pmusoke@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug, rwalusimbi@yalelo.ug'},
    {'WarehouseId' : 'Mukono', 'Email':'mukonoshop@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Mutungo', 'Email':'mutungostore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Nansana V2', 'Email':'nansanastore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Production', 'Email':'nbitsinze@yalelo.ug, mnanyibuka@yalelo.ug, ldemol@yalelo.ug, pomara@yalelo.ug, pmusoke@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Natete V2', 'Email':'nateetestore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Ntinda', 'Email':'ntindastore@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'KyebandoR', 'Email':'mbiryetega@yalelo.ug, pomara@yalelo.ug, pmusoke@yalelo.ug, bkabuye@yalelo.ug, bmbaraga@yalelo.ug'},
    {'WarehouseId' : 'HighValue', 'Email':'bkabuye@yalelo.ug, bmbaraga@yalelo.ug, anagudi@yalelo.ug, wekisa@yalelo.ug, dabigaba@yalelo.ug, ldemol@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Nyahuka', 'Email':'nyahukaborder@yalelo.ug, anagudi@yalelo.ug, wekisa@yalelo.ug, dabigaba@yalelo.ug, ldemol@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'Odramacaku', 'Email':'odramacakustore@yalelo.ug, anagudi@yalelo.ug, wekisa@yalelo.ug, dabigaba@yalelo.ug, ldemol@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug'},
    {'WarehouseId' : 'KyebandoDC', 'Email':'mbiryetega@yalelo.ug, nbitsinze@yalelo.ug, mnanyibuka@yalelo.ug, ldemol@yalelo.ug, pomara@yalelo.ug, pmusoke@yalelo.ug, bkabuye@yalelo.ug, bmbaraga@yalelo.ug, alakica@yalelo.ug, wekisa@yalelo.ug, pokuttu@yalelo.ug, rnabukeera@yalelo.ug, rwalusimbi@yalelo.ug'}
    ]


active_warehouse = pd.DataFrame(active_warehouse)

#%%

# //*[@id="email"]
#navigate the report and take screenshot
current_dir= "P:/Pertinent Files/Python/scripts/daily_dispatch_status/"
ss_name ='distribution.png'
ss_path=current_dir+ss_name
# ss_path = os.path.join(current_dir, "/", ss_name)
# print(f"This is the picture path: {ss_path}")

transfer_in_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/e4e07d5f73507b70b7bd"
transfer_out_url = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/74da8449-897f-467a-be26-56a067be3b0c/203e64e3a7c660b4d76a"
actual_received_stock = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/327e7fb0-8c3c-4596-bdd3-98e41603c2a6/a25014801c32e2406317"

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

    WebDriverWait(browser, 5*60).until(
            EC.presence_of_element_located((By.XPATH,'//*[@data-testid="visual-more-options-btn" and @class="vcMenuBtn" and @aria-expanded="false" and @aria-label="More options"]'))
        ).click()

    #Click Export data
    WebDriverWait(browser, 5*60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@title="Export data"]'))
    ).click()

    #Click Export Icon
    WebDriverWait(browser, 5*60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@aria-label="Export"]'))
    ).click()

    # Wait for the file to download (this might need adjustments based on download time)
    WebDriverWait(browser, 5*60).until(
        lambda driver: len(glob.glob("C:/Users/Administrator/Downloads/data*.xlsx")) > len(file_path)
    )
    time.sleep(2)
    return print("file download completed!!!")

#%%
# read files
export(transfer_in_url)

#%%
tr_in = pd.read_excel("C:/Users/Administrator/Downloads/data.xlsx")

#%%
df_in = pd.pivot_table(tr_in,index=["Date", "ReceivingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])
df_in['Total'] = df_in.sum(axis=1, skipna=True)

#%%
#DC transfers from production
df_in_dc_pdn = pd.pivot_table(tr_in[(tr_in['Received (kg)'] > 15) 
                                    & ((tr_in['ShippingWarehouseId'] == 'Production')|(tr_in['ShippingWarehouseId'] == 'HighValue')) 
                                    & (tr_in['ReceivingWarehouseId'] == 'KyebandoDC')],
                              index=["Date", "ReceivingWarehouseId","ShippingWarehouseId"], 
                              columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])

df_in_dc_pdn['Total Received'] = df_in_dc_pdn.sum(axis=1, skipna=True)
# print(df_in_dc_pdn.head())
#%%
#DC transfers from other warehouse
df_in_dc_wh = pd.pivot_table(tr_in[(tr_in['Received (kg)'] > 15) 
                                   & (((tr_in['ShippingWarehouseId'] != 'Production') & (tr_in['ShippingWarehouseId'] != 'HighValue')) ) 
                                   & (tr_in['ReceivingWarehouseId'] == 'KyebandoDC')],
                                    index=["Date", "ReceivingWarehouseId","ShippingWarehouseId"], 
                                    columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index(names=['Date','ReceivingWarehouseId'])
df_in_dc_wh['Total Received Frm WHS'] = df_in_dc_wh.sum(axis=1, skipna=True)
# 
# print(df_in_dc_wh.head())
#%%
#Driploss transfers from other warehouse
df_dl_dc = pd.pivot_table(tr_in[(tr_in['Received (kg)'] < 15) 
                                & ((tr_in['ShippingWarehouseId'] != 'Production')|(tr_in['ShippingWarehouseId'] != 'HighValue')) 
                                & (tr_in['ReceivingWarehouseId'] == 'KyebandoDC')],
                                index=["Date", "ReceivingWarehouseId","ShippingWarehouseId"], 
                                columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')
df_dl_dc['Driploss From Shops'] = df_dl_dc.sum(axis=1, skipna=True)

# print(df_dl_dc.head())
#%%
export(transfer_out_url)
tr_out = pd.read_excel("C:/Users/Administrator/Downloads/data.xlsx")
df_out = pd.pivot_table(tr_out,index=["Date","ReceivingWarehouseId","ShippingWarehouseId"], columns=["SKU","ProductSizeId"], values="Received (kg)", aggfunc='sum')#.reset_index()
df_out['Total Shipped Out'] = df_out.sum(axis=1, skipna=True)
#%%
export(actual_received_stock)
net_received_stock = pd.read_excel("C:/Users/Administrator/Downloads/data.xlsx")
net_received_stock = net_received_stock.iloc[:-3,:]
# print(net_received_stock)

#%%
# close
browser.quit()
#%%

#%%
# log-in to email server
"""
# ######################################################################
# # Email With Attachments Python Script
# # 
# ######################################################################
# """
# 
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
# 
cc_list = ["pokuttu@yalelo.ug, rnabukeera@yalelo.ug"]
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
TIE_server.login(email_from, pbi_pass_email)
print("Succesfully connected to server")
#%%
for i,e in zip(active_warehouse['WarehouseId'].to_list(), active_warehouse['Email'].to_list()):
    email_list = [email.strip() for email in e.split(',')] 
    # cc_list =  ["pokuttu@yalelo.ug","rnabukeera@yalelo.ug"]
    try:
        act_received  = net_received_stock[net_received_stock["WarehouseId"]==i]
        print(act_received.shape)
        
    except:
        None
    try:
        df_in_w = df_in.loc[(slice(None), i),:].dropna(axis=1, how='all')
        df_in_w.loc['Sub Total', :] = df_in_w.sum().values
    except:
        None
    try:
        df_out_w = df_out.loc[slice(None), slice(None), i].dropna(axis=1, how='all')
        df_out_w.loc['Sub Total', :] = df_out_w.sum().values
    except:
        None
    try:
        df_in_dc_pdn = df_in_dc_pdn.loc[slice(None), i, slice(None)].dropna(axis=1, how='all')
        df_in_dc_pdn.loc['Sub Total', :] = df_in_dc_pdn.sum().values
        print(f"kyebando data: {df_in_dc_pdn.shape}")
        # time.sleep(100)
    except:
        print("Kyebando has no data")
        # time.sleep(100)
    try:
        df_in_dc_wh = df_in_dc_wh.loc[slice(None), i, slice(None)].dropna(axis=1, how='all')
        df_in_dc_wh.loc['Sub Total', :] = df_in_dc_wh.sum().values
    except:
        None
# 
    try:
        df_dl_dc = df_dl_dc.loc[slice(None), i, slice(None)].dropna(axis=1, how='all')
        df_dl_dc.loc['Sub Total', :] = df_dl_dc.sum().values
    except:
        None

    if i=='KyebandoDC':
        with pd.ExcelWriter(f"{i} MTD Stock.xlsx") as writer:
                # Write df_in_w to the first sheet
                try:
                    df_in_dc_pdn.to_excel(writer, sheet_name=f'{i} Received(Frm Pdn)')
                except:
                    None
                # Write df_in_w to the first sheet
                try:
                    df_dl_dc.to_excel(writer, sheet_name=f'{i} Driploss(Frm Shops)')
                except:
                    None
                # Write df_in_w to the first sheet
                try:
                    df_in_dc_wh.to_excel(writer, sheet_name=f'{i} Received(WHS)')
                except:
                    None
                # Write df_out_w to the second sheet
                try:
                   df_out_w.to_excel(writer, sheet_name=f'{i} Stock Shipped Out')
                except:
                    None
    else:
        with pd.ExcelWriter(f"{i} MTD Stock.xlsx") as writer:
            # Write df_in_w to the first sheet
            try:
                act_received.to_excel(writer, sheet_name=f'{i} Net Stock Received')
            except:
                None
            try:
                df_in_w.to_excel(writer, sheet_name=f'{i} Gross Stock Received')
            except:
                None
            # Write df_out_w to the second sheet
            try:
                df_out_w.to_excel(writer, sheet_name=f'{i} Stock Shipped Out')
            except:
                None
#    name the email subject
# 
    subject = f"{i} Transfers Report [Month To Date]"

    # Prepare the body of the email
    if i == 'KyebandoDC':
        body = f"""
        Hello {i} Centre,

        Please find your MTD transfers report. There are two sheets (transfers in and transfers out), ensure to check of both.
        The Total Received - Total Shipped Stock = Net Received (shared daily on whatsapp)

        In case of any issues, please reply to pokuttu@yalelo.ug for followup.

        Regards,
        Pole
        """
    elif i in {'Malaba','Busia','Nyahuka','Mpondwe'}:
        body = f"""
        Hello {i} Border Point,

        Please find your MTD transfers report. There are two sheets (transfers in and transfers out), ensure to check of both.
        The Total Received - Total Shipped Stock = Net Received (shared daily on whatsapp)

        In case of any issues, please reply to pokuttu@yalelo.ug and pomara@yalelo.ug in copy for followup.

        Regards,
        Pole
        """
    else:
        body = f"""
        Hello {i} Store,

        Please find your MTD transfers report. There are two sheets (transfers in and transfers out), ensure to check of both.
        The Total Received - Total Shipped Stock = Net Received (shared daily on whatsapp)

        In case of any issues, please reply to pokuttu@yalelo.ug and mbiryetega@yalelo.ug in copy for followup.

        Regards,
        Pole
        """
    # Make a MIME object to define parts of the email
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['Subject'] = subject
    msg['To'] = ', '.join(email_list)
    msg['Cc'] = ', '.join(cc_list)
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
import glob
import os

for k in glob.glob("P:/Pertinent Files/"+ "*MTD Stock*" + "*.xlsx"):
    os.remove(k)
try:
    for g in glob.glob("C:/Users/Administrator/Documents/Python_Automations/"+ "*MTD Stock*" + "*.xlsx"):
        os.remove(g)
except:
    None
try:
    for g in glob.glob("C:/Users/Administrator/Documents/Python_Automations/Finance/"+ "*MTD Stock*" + "*.xlsx"):
        os.remove(g)
except:
    None

print("All files removed from repository")

# %%
