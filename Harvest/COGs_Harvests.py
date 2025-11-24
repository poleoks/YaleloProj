import sys
#file path
file_path_r = 'C:/Users/Administrator/Documents/Python_Automations/'


# file_path_r = 'D:/YU/ScriptCodez/'

sys.path.append(file_path_r)
from powerbi_sign_in_file import *
# download_path = 'C:/Users/Pole Okuttu/Downloads/'
download_path = 'C:/Users/Administrator/Downloads/'
print("Modules imported successfully")
#%%
attachments_ = "C:/Users/Administrator/Downloads/data.xlsx"

subject = "Harvest COGS Report"
body = "Please find attached the Harvest COGS report."
to = "pokuttu@yalelo.ug"

time.sleep(5)
#%%
for h in attachments_:
    try:
        os.remove(h)
        print(f"{h} removed")
    except FileNotFoundError:
        print(f"{h} not found, skipping removal.")

pbi_export('https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/7095f1e2-0598-4ec9-ba08-cc4f2cd2f9c6/818e08a4e3e8b2785e2c', f"{download_path}")

print("Download complete...")
dd=pd.read_excel(attachments_)
dd.to_excel("HarvestCOGs.xlsx", index=False)
from gmail_sender import *

gmail_function(to,subject, body, "HarvestCOGs.xlsx")
time.sleep(5)
