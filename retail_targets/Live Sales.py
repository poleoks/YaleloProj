import sys
save_dir = "C:/Users/Administrator/Documents/Python_Automations/"
sys.path.append(f'{save_dir}')
# sys.path.append('D:/YU/ScriptCodez/')
from powerbi_sign_in_file import *
from whatsapp_file_sign_in import *
from datetime import datetime#, timedelta
#import calendar

# NEW: explicit imports so the script doesn't depend on the star imports above
import os, glob, time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
)

DOWNLOADS = "C:/Users/Administrator/Downloads/"
TX_PATTERN = f"{DOWNLOADS}Transactions*.xlsx"

#%%
# NEW: helper functions
def snap(name):
    """Save a screenshot + print the URL, for debugging failures."""
    path = f"{save_dir}retail_targets/{name}.png"
    browser.save_screenshot(path)
    print(f"Screenshot saved: {path} | URL: {browser.current_url}")

def wait_overlay(timeout=30):
    """Wait for D365's blocking overlay to disappear."""
    WebDriverWait(browser, timeout).until(
        EC.invisibility_of_element_located((By.ID, "ShellBlockingDiv"))
    )

def safe_click(xpath, timeout=15):
    """Wait for the overlay to clear, then click. Retries once via JS if intercepted."""
    wait_overlay()
    el = WebDriverWait(browser, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    try:
        el.click()
    except ElementClickInterceptedException:
        time.sleep(2)
        wait_overlay()
        el = browser.find_element(By.XPATH, xpath)  # re-find in case it went stale
        browser.execute_script("arguments[0].click();", el)

#%%
# CHANGED: clean up old files separately so one failure doesn't skip the rest
i = 0
for pattern in [f"{save_dir}retail_targets/*.xlsx", TX_PATTERN, f"{save_dir}retail_targets/*.png"]:
    for h in glob.glob(pattern):
        try:
            os.remove(h)
            i += 1
            print(f"{h} removed! {i}")
        except Exception as e:
            print(f"Could not remove {h}: {e}")
if i == 0:
    print("No temp files to remove, proceed!")

#%%
# CHANGED: start_day removed (it was never used; the report only needs today)
today = datetime.today()
now = datetime.now().strftime("%Y-%m-%d %H:%M")
end_day = today.strftime('%m/%d/%Y')

print("Filtering for:", end_day)
time.sleep(2)

#%%
# CHANGED: removed browser.delete_all_cookies() (it wiped the login session)
browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=RetailRetailStoreTransactionTable')

try:
    WebDriverWait(browser, 90).until(
        EC.element_to_be_clickable((By.XPATH, '//*[contains(@id,"Grid_3") and @aria-label="Grid options"]'))
    ).click()
except TimeoutException:
    snap("grid_options_timeout")
    raise

print("more labels clicked")
time.sleep(3)

# Insert column
clk = WebDriverWait(browser, 15).until(
    EC.presence_of_element_located((By.XPATH, '//*[@aria-label="Insert columns..."]'))
)
print("click!")
time.sleep(2)
clk.click()
print("insert cols clicked")

# CHANGED: per-column try/except (no more bare "except: pass"), plus overlay waits
for cols in ["Number of products", "Transaction time", "Gross amount", "Warehouse"]:
    try:
        filter = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.XPATH, '(//input[@aria-label="Filter"])[2]'))
        )
        filter.clear()
        time.sleep(1)

        filter.send_keys(cols)
        print(f"{cols} pasted")
        time.sleep(2)

        safe_click('//*[@class="quickFilter-listFieldName" and contains(text(), "Field")]', 10)
        time.sleep(2)

        safe_click('//*[contains(@class,"dyn-checkbox-span")]', 10)
        wait_overlay(15)
        time.sleep(1)
        print(f"{cols} added")
    except Exception as e:
        print(f"{cols} FAILED: {type(e).__name__}")
        snap(f"fail_{cols.replace(' ', '_')}")

# CHANGED: single Update click (it was duplicated), now waits for the overlay
try:
    safe_click('//*[@class="button-label" and contains(text(),"Update")]', 15)
    print("update clicked")
except Exception as e:
    snap("update_failed")
    raise

#%%
# Financial Date Filter (via the grid quick filter)
try:
    wait_overlay(30)
    time.sleep(3)

    # The visible quick-filter box above the grid
    qf = next(i for i in browser.find_elements(By.XPATH, '//input[@name="QuickFilterControl_Input"]')
              if i.is_displayed())
    qf.click()
    qf.clear()
    qf.send_keys(end_day)
    time.sleep(2)

    # Debug: show which fields the dropdown offers
    opts = [o.text for o in browser.find_elements(By.XPATH, '//*[contains(@class,"quickFilter-listFieldName")]')
            if o.is_displayed()]
    print("Quick filter options:", opts)

    # Pick the "Date" field (exact match, so it doesn't hit e.g. "Statement date")
    safe_click('//*[contains(@class,"quickFilter-listFieldName") and normalize-space()="Date"]', 15)

    wait_overlay(60)
    time.sleep(3)
    snap("after_date_filter")
except (TimeoutException, StopIteration):
    snap("date_filter_failed")
    raise

print("end date filtered")
time.sleep(5)
wait_overlay(60)

#%%
# Export
WebDriverWait(browser, 15).until(
    EC.presence_of_element_located(
        (By.XPATH, '//*[@class="button-commandRing MicrosoftOffice-symbol"]')
    )
).click()
print('click office')
time.sleep(2)

# click export excel
WebDriverWait(browser, 15).until(
    EC.presence_of_element_located((By.XPATH, '(//*[text()="Transactions"])[2]'))
).click()
print("export menu")

# CHANGED: remember what's already in Downloads so we only pick up the NEW file
files_before = set(glob.glob(TX_PATTERN))

# click download button
WebDriverWait(browser, 15).until(
    EC.presence_of_element_located((By.XPATH, '//*[@data-dyn-controlname="DownloadButton"]'))
).click()
print("export started")

# CHANGED: wait for a new file AND for Chrome to finish writing it
WebDriverWait(browser, 45 * 60).until(
    lambda driver: bool(set(glob.glob(TX_PATTERN)) - files_before)
    and not glob.glob(f"{DOWNLOADS}*.crdownload")
)
time.sleep(2)
print("export completed")

# CHANGED: read the newest new file instead of whichever was last in the loop
new_files = set(glob.glob(TX_PATTERN)) - files_before
latest_file = max(new_files, key=os.path.getmtime)
print("Reading:", latest_file)
df = pd.read_excel(latest_file)

print(df.shape)

try:
    kill_browser("chrome")
    print("Running browser killed")
except:
    print("No Running browser, proceed")
    pass

#%%
import matplotlib.pyplot as plt
print(df.columns, df.head())
border_depos = ['Nyahuka', 'Busia', 'Odramacaku', 'Mpondwe', 'Elegu', 'Malaba', 'Kisoro', 'Gulu','Mbale','Jinja V3','Lira','Mbarara']
now = datetime.now().strftime("%d-%b-%Y %I:%M %p")
print(f"{now} - Live Sales Chart Generation Started")

df = df[
    (df['Number of products'] > 0) &
    (df['Warehouse'] != "Farmgate") &
    (df['Warehouse'] != "FarmgatUGX") &
    (df['Receipt number'].str.len() > 0)
]
df_border = df[df['Warehouse'].isin(border_depos)]
df_nonborder = df[~df['Warehouse'].isin(border_depos)]

# CHANGED: removed unused overall_df_bdr / overall_df_nbdr

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
# Send whatsapp message
groups_t = ['YU Retail Team','YU Rest of Country Sales Team']
files_t = ["nonborder_warehouse_value_chart.png","border_warehouse_value_chart.png"]
messages_t = [f"Latest Sales GKMA/RC Shops: {nonborder_total:.2f}",f"Latest Sales BDR: {border_total:.2f}"]

whatsapp_share(groups_t, messages_t, files_t, save_dir, Justine)
time.sleep(5)

#%%
for i in glob.glob(f"{save_dir}*.png"):
    os.remove(i)
    print(f"{i} removed successfully!")
print("All temp files removed successfully!")
time.sleep(3)
kill_browser("chrome")
print("Running browser killed successfully!")