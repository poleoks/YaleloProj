import os
print(os.getcwd())
from shareplum import Site, Office365
from shareplum.site import Version
import json
from creds import *
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