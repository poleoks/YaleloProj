#%%
import os
import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/')
from credentials import  *
from harvest_data_logmanager1 import *


#%%
#Write data
file_path='C:/Users/Administrator/Documents/Python_Automations/Harvest/YU-Harvest Data Logger.xlsx'

# with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
#     df.to_excel(writer, index=False, sheet_name='Sheet1')
#     workbook = writer.book
#     worksheet = writer.sheets['Sheet1']
#     table_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
#     header_format = workbook.add_format({'bold': True, 'bg_color': '#B0E0E6', 'border': 1})
#     worksheet.add_table(0, 0, df.shape[0], df.shape[1] - 1, {'style': 'Table Style Light 9'})
#     for col_num, value in enumerate(df.columns.values):
#         worksheet.write(0, col_num, value, header_format)

# df['datetime'] = pd.to_datetime(df['datetime'])
# df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')


# hv = pd.ExcelWriter(file_path)
df.to_excel(file_path, index=False, sheet_name='Sheet1',startrow=0,header=True)


#%%
"""
######################################################################
# Email With Attachments Python Script
######################################################################
"""

# name the email subject
subject = f"Harvest Report"

# import sys 
new_path='C:/Users/Administrator/Documents/Python_Automations/Harvest/'
# sys.path.insert(0, new_path)
os.chdir(new_path)
# Prepare the body of the email
body = """
Hello Team,

Please find Harvest Report attached.

Regards,
Audit Team
"""

filename = "YU-Harvest Data Logger.xlsx"

# Run the function
from gmail_sender import *
gmail_function('pokuttu@yalelo.ug',subject,body,file_path)

try:
    os.remove(file_path)
    # for path in  in glob.ig
    print("File deleted from repo")
except:
    print("File Does not Exist")
# %%
