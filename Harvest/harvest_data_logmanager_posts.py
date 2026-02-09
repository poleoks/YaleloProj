#%%
import sys
sys.path.append("C:/Users/Administrator/Documents/Python_Automations/")
import warnings
warnings.filterwarnings('ignore')
import mysql.connector
import pandas as pd
from credentials import *
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
    host=host_pin,
    user=user_pin,
    password=password_pin,
    database=database_pin
)

my_cursor=conn.cursor()

# SQL query to select top 5 rows
sql = "SELECT * FROM productionlogsheet lg where DATE(lg.datetime) = CURRENT_DATE"#-1"
# sql  = "SHOW DATABASES"
my_cursor.execute(sql)

# Fetch the results
df_mysql = my_cursor.fetchall()
print(df_mysql)
col_names=[i[0] for i in my_cursor.description]
# Print the df_mysql
# for row in df_mysql:
#     print(row)
harvest_data = pd.DataFrame(df_mysql, columns=col_names).set_index('id')
harvest_data['batch_number'] = harvest_data['batch_number'].str.strip()
df=harvest_data.astype('str').copy().sort_values(by='datetime')
print(df.tail())

# Total sum of netweight
df['datetime'] = pd.to_datetime(df['datetime'])
total_weight = df['netweight'].astype(float).sum() / 1000  # Convert to tons

# Total duration in hours
time_span = df['datetime'].max() - df['datetime'].min()
total_hours = time_span.total_seconds() / 3600 

# Calculation
if total_hours > 0:
    avg_weight_per_hour = total_weight / (total_hours)
else:
    avg_weight_per_hour = 0

print(f"Total Weight: {total_weight:.2f} kg")
print(f"Total Hours: {total_hours:.2f}")
print(f"Ton/Hr: {avg_weight_per_hour:.2f}")


#%%
# # check tables
# import mysql.connector

# # Execute the command
# my_cursor.execute("SHOW TABLES")

# # Fetch all results
# tables = my_cursor.fetchall()

# print("Tables in the database:")
# for table in tables:
#     my_cursor.execute(f"SELECT COUNT(*) FROM {table[0]}")
#     count = my_cursor.fetchone()[0]
#     print(f"- {table[0]}: {count} records")
#     print(f"- {table[0]}")