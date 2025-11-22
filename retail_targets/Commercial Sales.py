#%%
# storage local based on time
import sys
save_dir= "C:/Users/Administrator/Documents/Python_Automations/"
sys.path.append(save_dir)
from whatsapp_file_sign_in import *
from powerbi_sign_in_file import *

now = datetime.datetime.now()
noon = now.replace(hour=12, minute=0, second=0, microsecond=0)
yesterday = (now - datetime.timedelta(days=1)).date()
whatsapp_grp = ['YU Retail Team','Area Sales Leads','YU Rest of Country Sales Team','YU Rest of Country Sales Team',
                'YU Retail Team','Area Sales Leads','YU Sales Executives','YU Sales Executives']
if now < noon:
    pic_files = [
         f'{save_dir}retail.png',
         f'{save_dir}area_sales.png',
         f'{save_dir}rc_sales.png',
         f'{save_dir}risk_analysis.png',
         f'{save_dir}risk_analysis.png',
         f'{save_dir}risk_analysis.png',
         f'{save_dir}ooh_gkma.png',
         f'{save_dir}resellers_gkma.png'
         ]
    msg = [f"MTD GKMA Sales Report {yesterday}",
           f"MTD Area Sales Report {yesterday}",
           f"MTD BODR/RC Sales Report {yesterday}",
           f"MTD Banking Report {yesterday}, please submit banking by 12:59PM",
           f"MTD Banking Report {yesterday}, please submit banking by 12:59PM",
           f"MTD Banking Report {yesterday}, please submit banking by 12:59PM",
           f"MTD OOH Report {yesterday}",
           f"MTD Resellers Report {yesterday}"
           ]
    print(len(pic_files), now,"<",noon)
    print("The current time is before midday.")
else:
    pic_files = [
                f'{save_dir}risk_analysis.png',
                f'{save_dir}risk_analysis.png',
                f'{save_dir}risk_analysis.png',
                f'{save_dir}risk_analysis.png'
                ]
    msg = [
           f"MTD Banking Report {yesterday}",
           f"MTD Banking Report {yesterday}",
           f"MTD Banking Report {yesterday}"
           ]
    print(len(pic_files), now,">",noon)
    print("The current time is midday or later.")
#%%

retail_grp = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/c04614698037b6c54570"
area_sales_grp = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/d07599e95d4f8b1765bd"
rc_sales = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/a9b1381a5590a908d025"
risk_analysis = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/91687127-028f-4c1e-8ad9-3f783f724150/ReportSection7fb968f4825ec2e5889b"
ooh_gkma = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/53ec76f15a2501eece98"
resellers_gkma = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/be0d3b60c587d413ab0e"

if now < noon:
    files_t =['retail.png','area_sales.png','rc_sales.png','risk_analysis.png',
              'risk_analysis.png','risk_analysis.png','ooh_gkma.png','resellers_gkma.png']
    
    groups_t = ['YU Retail Team','Area Sales Leads','YU Rest of Country Sales Team',
                'YU Rest of Country Sales Team','Area Sales Leads','YU Retail Team',
                'YU Retail Team','YU Retail Team']
    
    messages_t = [f"GKMA Sales: {yesterday}",f"Area Sales: {yesterday}",f"RC/BDR Sales: {yesterday}",
                  f"MTD/Yesterday Banking: {yesterday}",f"MTD/Yesterday Banking: {yesterday}",
                  f"MTD/Yesterday Banking: {yesterday}",f"OOH Sales: {yesterday}",
                  f"Reseller Sales: {yesterday}"]

    pbi_sign_in(retail_grp)
    ss_path_ret = save_dir+'retail.png'
    browser.save_screenshot(ss_path_ret)
    print("Retail successfully taken screenshot!!!")
    time.sleep(5)
    ss_path_area = save_dir+'area_sales.png'
    pbi_sign_in(area_sales_grp)
    browser.save_screenshot(ss_path_area)
    print("Area sales successfully taken screenshot!!!")
    time.sleep(5)

    ss_path_rc = save_dir+'rc_sales.png'
    pbi_sign_in(rc_sales)
    browser.save_screenshot(ss_path_rc)
    print("RC successfully taken screenshot!!!")
    time.sleep(5)

    ss_path_risk = save_dir+'risk_analysis.png'
    pbi_sign_in(risk_analysis)
    browser.save_screenshot(ss_path_risk)
    print("RC successfully taken screenshot!!!")

    ss_path_ooh = save_dir+'ooh_gkma.png'
    pbi_sign_in(ooh_gkma)
    browser.save_screenshot(ss_path_ooh)
    print("OOH successfully taken screenshot!!!")
    
    ss_path_resl = save_dir+'resellers_gkma.png'
    pbi_sign_in(resellers_gkma)
    browser.save_screenshot(ss_path_resl)
    print("Resellers successfully taken screenshot!!!")
else:
    files_t =['risk_analysis.png','risk_analysis.png','risk_analysis.png']
    
    groups_t = ['YU Rest of Country Sales Team','Area Sales Leads','YU Retail Team']
    
    messages_t = [
                  f"MTD/Yesterday Banking: {yesterday}",f"MTD/Yesterday Banking: {yesterday}",
                  f"MTD/Yesterday Banking: {yesterday}"
                  ]
    
    ss_path_risk = save_dir+'risk_analysis.png'
    pbi_sign_in(risk_analysis)
    browser.save_screenshot(ss_path_risk)
    print("RC successfully taken screenshot!!!")
    time.sleep(5)

time.sleep(2)
browser.quit()

#%%
#send whatsapp
time.sleep(2)
whatsapp_share(groups_t, messages_t,files_t, save_dir, Pole)
#%%
#delete files
import os
import glob
for i in glob.glob(f"{save_dir}*.png"):
    os.remove(i)
    print(f"{i} removed successfully!")
print("All temp files removed successfully!")
time.sleep(3)