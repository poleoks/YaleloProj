from harvest_data_logmanager1 import harvest_data,pd
from sharepoint import SharePoint
from openpyxl import Workbook
import json

import copy 
#%%
# Get harvest_data from db
# harvest_data.reset_index(drop=True, inplace=True)

harvest_data.set_index('id', inplace=True)

mysql_df=harvest_data.tail(1).astype('str').copy().to_dict('records')
# # mysql_df={'data':harvest_data.tail(1).astype('str').copy()}

# #%%
# # get clients sharepoint list
# clients = SharePoint().connect_to_list(ls_name='MySQL Harvest Data')
# # SharePoint_df = pd.DataFrame(clients)
# # SharePoint_df=SharePoint_df.tail().copy()

# # print(SharePoint_df)
# print(mysql_df)
# print(clients)
# #%%
# # Update list items
# clients.UpdateListItems(data=mysql_df,kind='Update')

# print("Done Done!!")
# #%%
# # SharePoint().append_to_list(ls_name='MySQL Harvest Data', my_sql=mysql_df.tail(2))