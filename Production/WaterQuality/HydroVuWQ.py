#import modules
import selenium
import time
import os
import datetime
import time
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementClickInterceptedException
from dateutil.relativedelta import relativedelta

import os
import pandas as pd 
import sys
# #%%
from selenium.webdriver.common.keys import Keys
import sys
sys.path.append('C:/Users/Administrator/Documents/Python_Automations/Distribution')

from whatsapp_configuration import CHROME_PROFILE_PATH_HydroVuWQ
# from powerbi_sign_in_file import *
#%%

#%%
current_dir= "C:/Users/Administrator/Documents/Python_Automations/Production/WaterQuality"

# ss_name ='Feed.png'
ss_path=current_dir#+ss_name

sys.path.append("C:/Users/Administrator/Documents/Python_Automations/")

#INSTANTIATE WHATSAPP
#%%
Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)
Options.add_argument(CHROME_PROFILE_PATH_HydroVuWQ)

chrome_install = ChromeDriverManager().install()

folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

Service = webdriver.ChromeService(chromedriver_path)

browser=webdriver.Chrome(service=Service, options = Options)
report_url = "https://www.hydrovu.com/#/vusitu-data-files/list"

browser.get(report_url)
time.sleep(5)

browser.find_element(By.XPATH,'//*[@id="main-ui-view"]/isi-banner-ui-view-full-height/ui-view/isi-vs-data-files/div/div/isi-manage-files-dropdown/div/button').click()
time.sleep(15)