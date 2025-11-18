import pyperclip
save_dir= "C:/Users/Administrator/Documents/Python_Automations/"
import sys
sys.path.append(save_dir)
from credentials import *
from powerbi_sign_in_file import *
from whatsapp_file_sign_in import *
#%%
dayofw = datetime.datetime.today().weekday()
yesterday= datetime.datetime.today() - timedelta(days=1)

print(dayofw % 2 == 0 )
if dayofw % 2 == 0:
    #Mon,Wed,Fri,Sun
    pbi_sign_in("https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/1ff5a0ec-ff78-4749-9d6a-122ce99d8595/cb72c88225967453e9c4")
else:
    #Tue,Thur,Sat
    pbi_sign_in("https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/1ff5a0ec-ff78-4749-9d6a-122ce99d8595/ca701930773c46c23172")
current_dir= "C:/Users/Administrator/Documents/Python_Automations/Distribution/"

ss_name ='commercial_stockloss.png'
ss_path=save_dir+ss_name
browser.save_screenshot(ss_path)

#%%
#send whatsapp
files_t =['commercial_stockloss.png','commercial_stockloss.png','commercial_stockloss.png']
groups_t = ['YU Retail Team','YU Rest of Country Sales Team','Andrew Yk Enterprise']
messages_t = [f"Stock loss as at: {yesterday}",f"Stock loss as at: {yesterday}",f"Stock loss as at: {yesterday}"]

whatsapp_share(groups_t, messages_t,files_t, save_dir, Pole)
time.sleep(5)

#%%
import os
import glob
for i in glob.glob(f"{save_dir}*.png"):
    os.remove(i)
    print(f"{i} removed successfully!")
print("All temp files removed successfully!")
time.sleep(3)
