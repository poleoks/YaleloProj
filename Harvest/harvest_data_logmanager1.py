#%%
import warnings
warnings.filterwarnings('ignore')
import mysql.connector
import pandas as pd
from sharepoint import *
import time
import requests
from requests.auth import HTTPBasicAuth

import datetime
import re
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side
import glob
import os
    
    # "C:\Users\Administrator\Documents\Python_Automations\Harvest\excel_upload_harvest_logger.py"

#%%
file_path='C:/Users/Administrator/Documents/Python_Automations/Harvest/YU - Harvest Data Logger.csv'
fp='C:/Users/Administrator/Documents/Python_Automations/Harvest/YU - Harvest Data Logger.xlsx'


try:
    os.remove(file_path)
except:
    None
    # print('file not found')


try:
    os.remove(fp)
except:
    None
    # print('not found')
#print sharepoint data
#%%
# get mysql data
conn=mysql.connector.connect(
    host='127.0.0.1',
    user='powerbi',
    password='Y@l3l0@2023',
    # port=3306,
    database='logmanager1'
)

my_cursor=conn.cursor()

# SQL query to select top 5 rows
sql = "SELECT * FROM productionlogsheet where DATE(datetime) >= CURRENT_DATE - INTERVAL 6 MONTH"
# where datetime <= '2024-02-20'
# Execute the query
my_cursor.execute(sql)

# Fetch the results
df_mysql = my_cursor.fetchall()
# print(df_mysql)
col_names=[i[0] for i in my_cursor.description]
# Print the df_mysql
# for row in df_mysql:
#     print(row)
harvest_data = pd.DataFrame(df_mysql, columns=col_names).set_index('id')
harvest_data['batch_number'] = harvest_data['batch_number'].str.strip()
df=harvest_data.astype('str').copy()
df.datetime = pd.to_datetime(df.datetime).dt.date 

df2 = df[df.datetime ==max(df.datetime)]
print(f"{df2.columns,df.tail()}, total harvest: {df2.netweight.astype('float').sum()}")