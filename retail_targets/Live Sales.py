import sys
save_dir = "C:/Users/Administrator/Documents/Python_Automations/"
sys.path.append(f'{save_dir}')
# sys.path.append('D:/YU/ScriptCodez/')
from powerbi_sign_in_file import *
from whatsapp_file_sign_in import *
from datetime import datetime#, timedelta
#import calendar

#%%

i=0
for h in glob.glob(f"{save_dir}retail_targets/*.xlsx") or glob.glob("C:/Users/Administrator/Downloads/Transactions*.xlsx") or glob.glob("*.png"):
        i += 1
        os.remove(h)
        
        print(f"{h} removed! {i}")
#%%
today = datetime.today()
first_day = today.replace(day=1) - relativedelta(months=1)
now = datetime.now().strftime("%Y-%m-%d %H:%M")
# Calculate the last day of the month

start_day=first_day.strftime('%m/%d/%Y')
end_day=today.strftime('%m/%d/%Y')

print(start_day, end_day)
time.sleep(2)

#%%
browser.delete_all_cookies()
browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=RetailRetailStoreTransactionTable')


#more labels
WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[contains(@id,"Grid_3") and @aria-label="Grid options"]'))
                ).click()

print("more labels clicked")
time.sleep(3)
# Insert column
# try:
clk=WebDriverWait(browser,15).until(
                EC.presence_of_element_located((By.XPATH,'//*[@aria-label="Insert columns..."]'))
                )
print("click!")
time.sleep(2)
clk.click()
print("insert cols clicked")

#warehouse search
try:
    for cols in ["Number of products","Transaction time","Gross amount","Warehouse"]:     
        filter = WebDriverWait(browser,5).until(
                        EC.presence_of_element_located((By.XPATH,'(//input[@aria-label="Filter"])[2]'))
                        )
        filter.clear()
        time.sleep(1)

        filter.send_keys(cols)
        print(f"{cols} pasted")

        time.sleep(2)
        browser.find_element(By.XPATH,'//*[@class="quickFilter-listFieldName" and contains(text(), "Field")]').click()
        time.sleep(2)

        browser.find_element(By.XPATH,'//*[contains(@class,"dyn-checkbox-span")]').click()
        time.sleep(2)
except:
       pass


browser.find_element(By.XPATH,'//*[@class="button-label" and contains(text(),"Update")]').click()
#Financial Date Filter


# Wait for blocking overlay to disappear
WebDriverWait(browser, 30).until(
    EC.invisibility_of_element_located((By.ID, "ShellBlockingDiv"))
)

# Activate date filter
target = WebDriverWait(browser, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="RBOTransactionTable_transDate_3_0_header"]'))
)
target.click()

# Begin date
bgd = WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@name="FilterField_RBOTransactionTable_transDate_transDate_Input_0"]'))
                )
bgd.clear()
time.sleep(1)
bgd.send_keys(end_day)
print("start date filtered")
# Start date
std = WebDriverWait(browser,60).until(
                EC.presence_of_element_located((By.XPATH,'//*[@name="FilterField_RBOTransactionTable_transDate_transDate_Input_1"]'))
                )
std.clear()
time.sleep(1)
std.send_keys(end_day)

time.sleep(1)
browser.find_element(By.XPATH,'//*[@id="__RBOTransactionTable_transDate_ApplyFilters_label"]').click()

print("end date filtered")

time.sleep(5)
WebDriverWait(browser, 15).until(
    EC.presence_of_element_located(
        (By.XPATH,'//*[@class="button-commandRing MicrosoftOffice-symbol"]')
    )
).click()
print('click office')
time.sleep(2)
#click export excel
WebDriverWait(browser, 15).until(
    EC.presence_of_element_located(
        (By.XPATH,'(//*[text()="Transactions"])[2]')
    )
).click()

print("export menu")
file_path =[]
#click download button
WebDriverWait(browser, 15).until(
    EC.presence_of_element_located(
        (By.XPATH,'//*[@data-dyn-controlname="DownloadButton"]')
    )
).click()
print("export started")
WebDriverWait(browser, 45*60).until(
                    lambda driver: len(glob.glob("C:/Users/Administrator/Downloads/Transactions*.xlsx")) > len(file_path)
                    )

print("export completed")
for file in glob.glob("C:/Users/Administrator/Downloads/Transactions*.xlsx"):
                    df = pd.read_excel(file)


print(df.shape)

try:
    kill_browser("chrome")
    print("Running browser killed")
except:
    print("No Running browser, proceed")
    pass

import matplotlib.pyplot as plt
print(df.columns, df.head())
border_depos = ['Nyahuka', 'Busia', 'Odramacaku', 'Mpondwe', 'Elegu', 'Malaba', 'Kisoro']
now = datetime.now().strftime("%d-%b-%Y %I:%M %p")
print()
df = df[(df['Number of products'] > 0) & (df['Warehouse'] != "Farmgate") & (df['Warehouse'] != "FarmgatUGX") & (len(df['Receipt number'])>0)]

df_border = df[df['Warehouse'].isin(border_depos)]
df_nonborder = df[~df['Warehouse'].isin(border_depos)]

overall_df_bdr = round(df_border['Number of products'].sum(), 2)
overall_df_nbdr = round(df_nonborder['Number of products'].sum(), 2)

def plot_and_save(df_sub, title_suffix, save_path):
    df_sub = df_sub.copy()
    # Extract hour from transaction time
    df_sub['Hour'] = pd.to_datetime(df_sub['Transaction time']).dt.hour

    # Assign time period based on hour
    def get_time_period(hour):
        if  hour < 12:
            return 'Morning (Before Noon)'
        elif 12 <= hour < 16:
            return 'Afternoon (12pm-4pm)'
        elif 16 <= hour <= 20:
            return 'Evening (4pm-8pm)'
        else:
            return 'After 8pm'

    df_sub['Time_Period'] = df_sub['Hour'].apply(get_time_period)

    # Define colors for each time period
    color_map = {
        'Morning (Before Noon)': 'mediumaquamarine',
        'Afternoon (12pm-4pm)': 'lightskyblue',
        'Evening (4pm-8pm)': 'moccasin',
        'After 8pm': 'salmon'
    }

    # Group by warehouse and time period
    pivot_df = df_sub.pivot_table(
        values='Number of products',
        index='Warehouse',
        columns='Time_Period',
        aggfunc='sum',
        fill_value=0
    )

    # Ensure all time periods exist (even if zero sales)
    for period in color_map.keys():
        if period not in pivot_df.columns:
            pivot_df[period] = 0

    # Reorder columns in chronological order
    period_order = ['Morning (Before Noon)', 'Afternoon (12pm-4pm)', 
                    'Evening (4pm-8pm)', 'After 8pm']
    pivot_df = pivot_df[period_order]

    # Sort by total sales
    pivot_df['Total'] = pivot_df.sum(axis=1)
    pivot_df = pivot_df.sort_values('Total', ascending=True)
    pivot_df = pivot_df.drop('Total', axis=1)

    # Create stacked horizontal bar chart
    fig, ax = plt.subplots(figsize=(10, 6))

    left = pd.Series([0] * len(pivot_df), index=pivot_df.index)

    for period in period_order:
        ax.barh(pivot_df.index, pivot_df[period], left=left, 
                color=color_map[period], label=period)
        left += pivot_df[period]

    ax.set_title(f'Live Sales Update Today (kg) - {title_suffix}\n{now}')
    ax.set_xlabel('Total Sales (kg)')
    ax.set_ylabel('Store')
    ax.grid(axis='x', linestyle='--', alpha=0.6)
    ax.legend(loc='lower right', fontsize=9)

    # Add total value labels at the end of each bar
    for idx, warehouse in enumerate(pivot_df.index):
        total = pivot_df.loc[warehouse].sum()
        ax.text(total + 1, idx, f'{total:.2f}', 
                va='center', fontsize=8, color='grey')

    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.close()

    return df_sub['Number of products'].sum()


# Save Both Charts

os.makedirs(save_dir, exist_ok=True)

border_path = os.path.join(save_dir, "border_warehouse_value_chart.png")
nonborder_path = os.path.join(save_dir, "nonborder_warehouse_value_chart.png")

border_total = plot_and_save(df_border, "Border Depots", border_path)
nonborder_total = plot_and_save(df_nonborder, "GKMA & RC Stores", nonborder_path)

print(f"Border total: {border_total:.2f} kg")
print(f"Non-border total: {nonborder_total:.2f} kg")


#%%
#Send whatsapp message
groups_t = ['YU Retail Team','YU Rest of Country Sales Team']
files_t = ["nonborder_warehouse_value_chart.png","border_warehouse_value_chart.png"]
messages_t = [f"Latest Sales GKMA/RC Shops: {nonborder_total:.2f}",f"Latest Sales BDR: {border_total:.2f}"]

whatsapp_share(groups_t, messages_t,files_t, save_dir, Justine)
time.sleep(5)

#%%
import os
import glob
for i in glob.glob(f"{save_dir}*.png"):
    os.remove(i)
    print(f"{i} removed successfully!")
print("All temp files removed successfully!")
time.sleep(3)