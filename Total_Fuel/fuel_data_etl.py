import sys
#file path
file_path_r = 'C:/Users/Administrator/Documents/Python_Automations/'
download_path = 'C:/Users/Administrator/Downloads/'

# file_path_r = 'D:/YU/ScriptCodez/'
# download_path = 'C:/Users/Administrator/Downloads/'

sys.path.append(file_path_r)

#%%
from credentials import  *

#%%
Options=webdriver.ChromeOptions()
Options.headless = True  # Enable headless mode
Options.add_experimental_option("detach", True)
#INSTANTIATE WHATSAPP
# sys.path.insert(0,'C:/Users/Administrator/Documents/Python_Automations/Distribution/')
from chrome_configuration import CHROME_PATH
Options.add_argument(CHROME_PATH)
chrome_install = ChromeDriverManager().install()

folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

Service = webdriver.ChromeService(chromedriver_path)
browser=webdriver.Chrome(service=Service, options = Options)
#%%
for file in glob.glob(os.path.join(download_path, '*.csv')):
    os.remove(file)
    print(f'Removed: {file}')

today = datetime.datetime.today()
# set today as yesterday
# today = today - timedelta(days=1)
year_ = today.year
month_ = today.month

today_=today.strftime('%m/%d/%Y')
last_yr_=today - timedelta(days=365)
last_yr_=last_yr_.strftime('%m/%d/%Y')


#GET FUEL CONSUMPTION IN KENYA
# groups_path='P:/Pertinent Files/Python/scripts/groups_all.txt'
browser.get('https://www.mytotalfuelcard.com/')
# <<<<<<< HEAD
browser.delete_all_cookies()

#>#>#>#>#>#>#> c03a598df93b05e6595dc8f7aa81f3b56fa360c0
browser.maximize_window()
time.sleep(5)

#%%
WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH,"//button[@id='tarteaucitronAllDenied2']"))
        ).click()

# WebDriverWait(browser, 50).until(
#         EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[3]/button[2]"))
#     ).click()

log_name=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/aside/div[1]/div[2]/form/div[1]/input"))
    )
# log_name=browser.find_element(By.XPATH,"//input[@id='tb_user_id']")
log_name.send_keys(Login_ke)

log_name=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/aside/div[1]/div[2]/form/div[2]/input"))
    )
# log_name=browser.find_element(By.XPATH,"//input[@id='tb_user_id']")
log_name.send_keys(Password_ke)
browser.find_element('xpath','/html/body/div[3]/aside/div[1]/div[2]/form/a/div').click()

# %%
transaction_button=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/aside/div[1]/nav/ul/li[3]/a"))
    )
transaction_button.click()


customer_filter=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/section/div/section/div/table/tbody/tr[1]/td[2]/select"))
    )
customer_filter.click()


select_yu=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/section/div/section/div/table/tbody/tr[1]/td[2]/select/option[2]"))
    )
select_yu.click()

# <<<<<<< HEAD
bgd = WebDriverWait(browser, 50).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="dp_date_debut"]'))
)

bgd.clear()
time.sleep(1)
bgd.send_keys('01/01/2024')
bgd.send_keys(Keys.ENTER)
time.sleep(1)

edd = WebDriverWait(browser, 50).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="dp_date_fin"]'))
)

edd.clear()
time.sleep(1)
edd.send_keys(today_)
edd.send_keys(Keys.ENTER)
time.sleep(1)
#%%
# csv download
download_link=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="csvIcon"]'))
    )
browser.execute_script("arguments[0].scrollIntoView(true);", download_link)


#%%
# csv download
download_link=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/section/div/section/div/div[3]/div[2]"))
    )
#>#>#>#>#>#>#> c03a598df93b05e6595dc8f7aa81f3b56fa360c0
download_link.click()

#%%
# Wait for the download to complete (you may need to customize the waiting time)
max_wait_time_seconds = 60*1
current_wait_time = 0

# <<<<<<< HEAD
while not os.path.exists("c:/Users/Administrator/Downloads/Transactions.csv") and current_wait_time < max_wait_time_seconds:
    time.sleep(1)
    current_wait_time += 1
# Check if the file exists
ke_fuel=pd.read_csv("c:/Users/Administrator/Downloads/Transactions.csv",sep=";")
# =======
while not os.path.exists("C:/Users/Administrator/Downloads/Transactions.csv") and current_wait_time < max_wait_time_seconds:
    time.sleep(1)
    current_wait_time += 1
# Check if the file exists
ke_fuel=pd.read_csv("C:/Users/Administrator/Downloads/Transactions.csv",sep=";")
#>#>#>#>#>#>#> c03a598df93b05e6595dc8f7aa81f3b56fa360c0
print(ke_fuel.shape)
print(ke_fuel.head())
#%%
for file in glob.glob(os.path.join(download_path, '*.csv')):
    os.remove(file)
    print(f'Removed: {file}')
#GET FUEL CONSUMPTION IN UGANDA
# groups_path='P:/Pertinent Files/Python/scripts/groups_all.txt'
browser.get('https://www.mytotalfuelcard.com/')
#>#>#>#>#>#>#> c03a598df93b05e6595dc8f7aa81f3b56fa360c0
browser.maximize_window()
time.sleep(5)

#%%
# browser.find_element(By.XPATH,"//button[@id='tarteaucitronAllDenied2']").click()

log_name=WebDriverWait(browser, 5).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/aside/div[1]/div[2]/form/div[1]/input"))
    )
# log_name=browser.find_element(By.XPATH,"//input[@id='tb_user_id']")
log_name.send_keys(Login_ug)

log_name=WebDriverWait(browser, 5).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/aside/div[1]/div[2]/form/div[2]/input"))
    )
# log_name=browser.find_element(By.XPATH,"//input[@id='tb_user_id']")
log_name.send_keys(Password_ug)
# log_password=browser.find_element(By.XPATH,"//input[@id='tb_password']")
# log_password.send_keys("no48d012")
browser.find_element('xpath','/html/body/div[3]/aside/div[1]/div[2]/form/a/div').click()

# %%
transaction_button=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/aside/div[1]/nav/ul/li[3]/a"))
    )
transaction_button.click()


customer_filter=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/section/div/section/div/table/tbody/tr[1]/td[2]/select"))
    )
customer_filter.click()


select_yu=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/section/div/section/div/table/tbody/tr[1]/td[2]/select/option[2]"))
    )
select_yu.click()


#%%
# <<<<<<< HEAD
bgd = WebDriverWait(browser, 50).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="dp_date_debut"]'))
)

# csv download
bgd.clear()
time.sleep(1)
bgd.send_keys(last_yr_)
bgd.send_keys(Keys.ENTER)
time.sleep(1)

edd = WebDriverWait(browser, 50).until(
    EC.presence_of_element_located((By.XPATH,'//*[@id="dp_date_fin"]'))
)

edd.clear()
time.sleep(1)
edd.send_keys(today_)
edd.send_keys(Keys.ENTER)
time.sleep(5)
#%%
# csv download
download_link=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH,'//*[@id="csvIcon"]'))
    )

browser.execute_script("arguments[0].scrollIntoView(true);", download_link)
# =======
# csv download
download_link=WebDriverWait(browser, 50).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/section/div/section/div/div[3]/div[2]"))
    )
#>#>#>#>#>#>#> c03a598df93b05e6595dc8f7aa81f3b56fa360c0
download_link.click()

#%%
# Wait for the download to complete (you may need to customize the waiting time)
# <<<<<<< HEAD
max_wait_time_seconds = 60*5
current_wait_time = 0

while not os.path.exists("c:/Users/Administrator/Downloads/Transactions.csv") and current_wait_time < max_wait_time_seconds:
    time.sleep(1)
    current_wait_time += 1
# Check if the file exists
ug_fuel=pd.read_csv("c:/Users/Administrator/Downloads/Transactions.csv",sep=";")

#%%
for file in glob.glob(os.path.join(download_path, '*.csv')):
    os.remove(file)
    print(f'Removed: {file}')
browser.quit()
time.sleep(5)

#%%
# create excel file with two sheets
attachment_path_ = 'C:/Users/Administrator/Downloads/YU_Fuel_Automated.xlsx'
try:
    os.remove(attachment_path_)
    print(f'Removed existing attachment: {attachment_path_}')
except FileNotFoundError:
    print(f'No existing attachment to remove at: {attachment_path_}')
    
with pd.ExcelWriter(attachment_path_) as writer:
    ke_fuel.to_excel(writer, sheet_name='Kenya Fuel', index=False)
    ug_fuel.to_excel(writer, sheet_name='Uganda Fuel', index=False)
print(f"uganda fuel data shape: {ug_fuel.shape} - kenya fuel data shape: {ke_fuel.shape}")
#%%
# ======= send gmail
from gmail_sender import *

gmail_function('pokuttu@yalelo.ug','YU Fuel Data Downloaded',
               'Fuel data for Kenya and Uganda has been downloaded successfully.',
               attachment_path_)

#%%
time.sleep(30)
os.remove(attachment_path_)
print(f'Removed attachment: {attachment_path_}')