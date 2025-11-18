#%%
import sys
# dir_path2 = 'C:/Users/Administrator/Documents/Python_Automations/'
dir_path = 'D:/YU/ScriptCodez/'
sys.path.append(f'{dir_path}')
from powerbi_sign_in_file import *
from whatsapp_file_sign_in import *
#%%
my_date=datetime.datetime.today() 
yesterday= datetime.datetime.today() - timedelta(days=1)
print(yesterday)

day_of_the_week=calendar.day_name[my_date.weekday()]
day_of_the_week_num=datetime.datetime.today().weekday() #0 for Monday, 1 for Tuesday
month_start_date =my_date.replace(day=1).strftime('%m/%d/%Y')
month_end_date = my_date.strftime('%m/%d/%Y')
yesterday_date = yesterday.strftime('%m/%d/%Y')


print(yesterday_date)
week_start_date=my_date - timedelta(days=my_date.weekday())
week_end_date=week_start_date+timedelta(days=6)
sd_date=week_start_date-timedelta(days=7)
ed_date=week_end_date- timedelta(days=7)

sd_date=sd_date.strftime('%m/%d/%Y')
ed_date=ed_date.strftime('%m/%d/%Y')

week_start_date=week_start_date.strftime('%m/%d/%Y')
week_end_date=week_end_date.strftime('%m/%d/%Y')


print(f"Month start date: {sd_date}")
print(f"Month end date: {ed_date}")

print(f"week start date: {week_start_date}")
print(f"week end date: {week_end_date}")

current_time = str(datetime.datetime.now().time())

#%%
current_dir= f"{dir_path}"

# ss_name ='Feed.png'
ss_path=current_dir#+ss_name

report_url="https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/e33a4d76-14e2-475d-b035-62fb4a47e444/50e17e2aa4e3c53a2f4c"

# wait for report load

#%%
pbi_sign_in(repo_url=report_url)
##wait for a max of 5 mins until full load
for jj in ['BS','GA','GB','GC','GD','GE','GG','NA','NB','NC']:
    try:
        browser.find_element(By.XPATH, f"//*[@title='{jj}']").click()
        time.sleep(5)
        # Take a screenshot of the report and save it to a file
        browser.save_screenshot(f"{ss_path}/{jj}_Feed.png")
        print(f"{jj} screenshot taken!")
    except:
        pass

    time.sleep(2)
print("successfully taken screenshots!!!")

time.sleep(5)

#%%

#send whatsapp messages
groups_t = ['Grow Out feeding', 'Grow Out feeding', 'Grow Out feeding', 'Grow Out feeding', 'Grow Out feeding', 'Grow Out feeding','Nursery YU', 'Nursery YU', 'Nursery YU', 
    'Nursery YU','FEED LOGISTICS', 'FEED LOGISTICS', 'FEED LOGISTICS', 'FEED LOGISTICS', 'FEED LOGISTICS', 'FEED LOGISTICS','FEED LOGISTICS', 
    'FEED LOGISTICS', 'FEED LOGISTICS', 'FEED LOGISTICS']
messages_t = [
    'Platform: GA_Feed', 'Platform: GB_Feed', 'Platform: GC_Feed', 'Platform: GD_Feed', 'Platform: GE_Feed', 'Platform: GF_Feed',
    'Platform: NA_Feed', 'Platform: NB_Feed', 'Platform: NC_Feed', 'Platform: BS_Feed','Platform: GA_Feed', 'Platform: GB_Feed', 'Platform: GC_Feed', 
    'Platform: GD_Feed', 'Platform: GE_Feed', 'Platform: GF_Feed','Platform: NA_Feed', 'Platform: NB_Feed', 'Platform: NC_Feed', 'Platform: BS_Feed'
    ]
files_t = [
    'GA_Feed.png', 'GB_Feed.png', 'GC_Feed.png', 'GD_Feed.png', 'GE_Feed.png', 'GF_Feed.png','NA_Feed.png', 'NB_Feed.png', 'NC_Feed.png', 
    'BS_Feed.png','GA_Feed.png', 'GB_Feed.png', 'GC_Feed.png', 'GD_Feed.png', 'GE_Feed.png', 'GF_Feed.png','NA_Feed.png', 
    'NB_Feed.png', 'NC_Feed.png', 'BS_Feed.png'
    ]

whatsapp_share(groups_t, messages_t,files_t, dir_path, Pole)
time.sleep(5)


for i in (f"{dir_path}*_Feed.png"):
    os.remove(i)
    print(f"{i} removed successfully!")
print("All temp files removed successfully!")
time.sleep(3)