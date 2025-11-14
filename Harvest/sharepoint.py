
# print(os.getcwd())
from shareplum import Site, Office365
from shareplum.site import Version
import json
from creds import *
import time
# from config import *

USERNAME=user
PASSWORD = password
SHAREPOINT_URL = url 
SHAREPOINT_SITE = site

    
class SharePoint:
    def auth(self):
        self.authcookie = Office365(
            SHAREPOINT_URL,
            username=USERNAME,
            password=PASSWORD,
        ).GetCookies()
        self.site = Site(
            SHAREPOINT_SITE,
            version=Version.v365,
            authcookie=self.authcookie,
        )
        return self.site

    def connect_to_list(self, ls_name):
        self.auth_site = self.auth()

        list_data = self.auth_site.List(list_name=ls_name).GetListItems()

        return list_data
    def append_to_list(self,ls_name,my_sql):
        my_sql=[my_sql.to_dict()]
        print(my_sql)
        time.sleep(5)
        self.auth_site = self.auth()

        # Get the SharePoint list
        sp_list = self.auth_site.List(list_name=ls_name)

        # Append data to the list
        sp_list.UpdateListItems(data=my_sql,kind='Update')
        
        print("Well Well, Good Luck!!")


# Certainly! To add the append functionality to your SharePoint class, you can create a method named append_to_list that takes the list name (ls_name) and the data you want to append (my_sql). Here's how you can modify your class:

# python
# Copy code
# import os
# from shareplum import Site, Office365
# from shareplum.site import Version
# import json
# from creds import *

# USERNAME = user
# PASSWORD = password
# SHAREPOINT_URL = url
# SHAREPOINT_SITE = site

# class SharePoint:
#     def auth(self):
#         self.authcookie = Office365(
#             SHAREPOINT_URL,
#             username=USERNAME,
#             password=PASSWORD,
#         ).GetCookies()
#         self.site = Site(
#             SHAREPOINT_SITE,
#             version=Version.v365,
#             authcookie=self.authcookie,
#         )
#         return self.site

#     def connect_to_list(self, ls_name):
#         self.auth_site = self.auth()

#         list_data = self.auth_site.List(list_name=ls_name).GetListItems()

#         return list_data

#     def append_to_list(self, ls_name, my_sql):
#         self.auth_site = self.auth()

#         # Get the SharePoint list
#         sp_list = self.auth_site.List(list_name=ls_name)

#         # Append data to the list
#         sp_list.UpdateListItems(data=my_sql, kind='New', update_method='UpdateOrCreate')

# # # Example usage
# # # Initialize the SharePoint object
# # sp = SharePoint()

# # # Authenticate
# # sp.auth()

# # # Append data to the SharePoint list
# # data_to_append = [{'Field1': 'Value1', 'Field2': 'Value2'}, {'Field1': 'Value3', 'Field2': 'Value4'}]
# # sp.append_to_list('YourListName', data_to_append)
