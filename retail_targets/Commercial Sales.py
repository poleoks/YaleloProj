#%%
# storage local based on time
import datetime

now = datetime.datetime.now()
noon = now.replace(hour=12, minute=0, second=0, microsecond=0)
yesterday = (now - datetime.timedelta(days=1)).date()
whatsapp_grp = ['YU Retail Team','Area Sales Leads','YU Rest of Country Sales Team','YU Rest of Country Sales Team',
                'YU Retail Team','Area Sales Leads','YU Sales Executives','YU Sales Executives']
if now < noon:
    pic_files = [
         'E:/Python_Automations/Pole/Distribution/retail.png',
         'E:/Python_Automations/Pole/Distribution/area_sales.png',
         'E:/Python_Automations/Pole/Distribution/rc_sales.png',
         'E:/Python_Automations/Pole/Distribution/risk_analysis.png',
         'E:/Python_Automations/Pole/Distribution/risk_analysis.png',
         'E:/Python_Automations/Pole/Distribution/risk_analysis.png',
         'E:/Python_Automations/Pole/Distribution/ooh_gkma.png',
         'E:/Python_Automations/Pole/Distribution/resellers_gkma.png'
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
                'E:/Python_Automations/Pole/Distribution/risk_analysis.png',
                'E:/Python_Automations/Pole/Distribution/risk_analysis.png',
                'E:/Python_Automations/Pole/Distribution/risk_analysis.png',
                'E:/Python_Automations/Pole/Distribution/risk_analysis.png'
                ]
    msg = [
           f"MTD Banking Report {yesterday}",
           f"MTD Banking Report {yesterday}",
           f"MTD Banking Report {yesterday}"
           ]
    print(len(pic_files), now,">",noon)
    print("The current time is midday or later.")

#%%
from selenium.webdriver.common.keys import Keys
import sys
sys.path.append("E:/Python_Automations/")
from whatsapp_config_pole import CHROME_PROFILE_PATH_Pole

from powerbi_sign_in_file import *

#%%
current_dir= "E:/Python_Automations/Pole/Distribution/"


ss_path=current_dir#+ss_name





retail_grp = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/c04614698037b6c54570"
area_sales_grp = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/d07599e95d4f8b1765bd"
rc_sales = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/a9b1381a5590a908d025"
risk_analysis = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/91687127-028f-4c1e-8ad9-3f783f724150/ReportSection7fb968f4825ec2e5889b"
ooh_gkma = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/53ec76f15a2501eece98"
resellers_gkma = "https://app.powerbi.com/groups/6514fc4d-2ddc-4df0-8cd7-1a6a5f7deed8/reports/65cfcd8b-d7af-442e-8e76-cd729a944fb1/be0d3b60c587d413ab0e"


def url(report_url):
    browser.get(report_url)
    try:
          reset_button= WebDriverWait(browser, 60*5).until(
                EC.presence_of_element_located((By.XPATH,'//*[@data-testid="reset-to-default-btn"]'))
                )
          
          reset_button.click()
          ok_button = WebDriverWait(browser, 15).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="okButton"]'))
                )
          ok_button.click()
          print("reset activated")
    except:
          print("reset not found")
          pass
    
    print("Expanding View")
    # Click the expand button
    view_button= WebDriverWait(browser, 15).until(
          EC.presence_of_element_located((By.XPATH,'//*[@data-testid="app-bar-view-menu-btn"]'))
    )
    view_button.send_keys(Keys.CONTROL + Keys.SHIFT + 'f')
    
    print("Expanded view")

    time.sleep(10)

    # Take a screenshot of the report and save it to a file
    # time.sleep(5)
if now < noon:
    url(retail_grp)
    browser.save_screenshot(ss_path+'retail.png')
    print("Retail successfully taken screenshot!!!")

    url(area_sales_grp)
    browser.save_screenshot(ss_path+'area_sales.png')
    print("Area sales successfully taken screenshot!!!")

    url(rc_sales)
    browser.save_screenshot(ss_path+'rc_sales.png')
    print("RC successfully taken screenshot!!!")

    time.sleep(5)

    url(risk_analysis)
    browser.save_screenshot(ss_path+'risk_analysis.png')
    print("RC successfully taken screenshot!!!")

    url(ooh_gkma)
    browser.save_screenshot(ss_path+'ooh_gkma.png')
    print("OOH successfully taken screenshot!!!")

    url(resellers_gkma)
    browser.save_screenshot(ss_path+'resellers_gkma.png')
    print("Resellers successfully taken screenshot!!!")
else:
    url(risk_analysis)
    browser.save_screenshot(ss_path+'risk_analysis.png')
    print("RC successfully taken screenshot!!!")

time.sleep(5)

browser.quit()

time.sleep(5)

#whatsapp  
Options=webdriver.ChromeOptions()
Options.add_experimental_option("detach", True)
Options.add_argument(CHROME_PROFILE_PATH_Pole)
Options.headless = True  # Enable headless mode

chrome_install = ChromeDriverManager().install()

folder = os.path.dirname(chrome_install)
chromedriver_path = os.path.join(folder, "chromedriver.exe")

Service = webdriver.ChromeService(chromedriver_path)

browser=webdriver.Chrome(service=Service, options = Options)
  
browser.get('https://web.whatsapp.com/')


for group,pfile,msg in zip(whatsapp_grp,pic_files,msg):
        # ss_path = '{group}'.format(group)
        search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
        search_box = WebDriverWait(browser, 500).until(
                EC.presence_of_element_located((By.XPATH, search_xpath))
        )
        search_box.clear()
        time.sleep(3)
        pyperclip.copy(group)
        search_box.send_keys(Keys.SHIFT, Keys.INSERT) # Keys.CONTROL + "v" for windows users
        time.sleep(2)

        group_xpath = f'//*[@title="{group}"]'
        group_title = WebDriverWait(browser, 500).until(
                EC.presence_of_element_located((By.XPATH, group_xpath))
        )
        # group_title = browser.find_elements(by=By.XPATH, value=group_xpath)
        # Test 3

        browser.execute_script("arguments[0].scrollIntoView();", group_title)
        group_title.click()
        time.sleep(5)

        attachment_box = browser.find_element(by=By.XPATH, value="//span[@data-icon='plus-rounded']")
        attachment_box.click()
        time.sleep(3)

        image_box =browser.find_element(by=By.XPATH, value='//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_box.send_keys(pfile)
        time.sleep(3)

        #%%
        txt_xpath = '//div[@contenteditable="true"][@role="textbox"]'
        txt_box = browser.find_element(by=By.XPATH, value=txt_xpath)

        # pyperclip.copy("Hi Team, appologies there was a glitch in the earlier sent report. Here is the updated on")
        pyperclip.copy(f'{msg}')
        txt_box.send_keys(Keys.SHIFT, Keys.INSERT)
        #send message
        txt_box.send_keys(Keys.ENTER)
        time.sleep(30)
        print("image sent via whatsapp....browser exiting!!!")


browser.quit()
time.sleep(2)