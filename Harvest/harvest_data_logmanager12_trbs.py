#%%
import warnings
warnings.filterwarnings('ignore')
import mysql.connector
import pandas as pd
from sharepoint import *
import time
import requests
from requests.auth import HTTPBasicAuth
# from excel_upload_harvest import *
# from projects import df_sharepoint
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
    print('file not found')


try:
    os.remove(fp)
except:
    print('not found')
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
sql = "SELECT * FROM productionlogsheet"
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
df=harvest_data.astype('str').copy()

print(df.tail())
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