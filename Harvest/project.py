from sharepoint import SharePoint
from openpyxl import Workbook
import pandas as pd
import json

# get clients sharepoint list
clients = SharePoint().connect_to_list(ls_name='YU Daily - Harvesting and Processing Packing')
df_sharepoint = pd.DataFrame(clients)
print(df_sharepoint.shape)

#get mysql data
