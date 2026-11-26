import sys
import random

save_dir = "C:/Users/Administrator/Documents/Python_Automations/"
# save_dir = "D:/YU/ScriptCodez/"

sys.path.append(f'{save_dir}')
from credentials import *   

from powerbi_sign_in_file import *
from datetime import datetime#, timedelta
from selenium.common.exceptions import TimeoutException


date_today=datetime.today().day
month_num=datetime.today().month
year=datetime.today().year
# Format the date as "mm/dd/yyyy" using strftime()
formatted_date = datetime.today().strftime('%m/%d/%Y')
end_date= datetime.today()

start_date= datetime.today() - timedelta(days=60)
start_date= start_date.strftime('%Y-%m-%d')
print(formatted_date)
# file_location=f"Customer Credit Report - {end_date}.xlsx"
download_path="C:/Users/Administrator/Downloads"
download_address = os.path.join(download_path, "Customer account statement*.pdf")

for d in glob.glob(download_address):
    if os.path.exists(d):
        os.remove(d)
        print(f"Deleted file: {d}")
    else:
        print(f"File not found: {d}")

try:
    browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=CustTableListPage')
    print("Browser initialized successfully")
    time.sleep(5)
    browser.maximize_window()
except:
    print("Error occurred while initializing the browser")


customer_list = {
    # "Bertha's Kitchen " : "berthamusinguzi@gmail.com",
    # "Oakridge catering  and hospitality " : "oaksridge.hosp@gmail.com",
    # "Latitude Zero Degrees Limited" : "0accounts2@thelatitudehotels.com",
    # "Sokony Africa Limited" : "accounts@qualityhill.biz,eowora@lepetitvillage.net,pnamususwa@lepetitvillage.net",
    "Cakely (U) Limited" : "Cafecakely@gmail.com",
    "Mapenzi Grill" : "Jeraldoboth14@gmail.com",
    "Extreme park Busiika" : "Accounts@extremeadventures.co.ug",
    "Villa Kololo" : "finance.villakololo@gmail.com",
    "Great Habib " : "minahnabs@gmail.com",
    "Kampala Serena" : "jeanne.ahairwe@serenahotels.com",
    "D A Fish and More Freshness Uganda" : "kiyingirichard32@gmail.com",
    "Rhino Restaurant " : "Info@comforthotel.co.ug",
    "Hotel la Vena " : "lavenahotel81@gmail.com",
    "Kalanoga resort " : "reservations@kalanoga.com",
    "Forest Park resort " : "forestparkresort@gmail.com",
    "Natures green resort " : "ybirigwa@yahoo.com",
    "Caramel café " : "caramelacc2018@gmail.com",
    "Zal Psalms resort" : "zalpalms@gmail.com",
    "Fresh Kyenyanja" : "benardnsengiyunva@gmail.com",
    "Jinja Sailing Club" : "accounts@jinjasailingclub.com",
    "MADA HOLDINGS ( JINJA NILE RESORT" : "fcug@madahotels.com"
    }
for customer, gmail_ac in customer_list.items():  
    time.sleep(3)
    browser.get('https://fw-d365-prod.operations.dynamics.com/?cmp=yu&mi=CustTableListPage')

    time.sleep(3)
    Customer_filt=WebDriverWait(browser, 20).until(
         EC.presence_of_element_located((By.XPATH, '//*[@name="QuickFilterControl_Input"]'))
    )
    Customer_filt.click()
    Customer_filt.clear()
    Customer_filt.send_keys(customer)
    print(f"Searching for {customer}...")

    xpath = '//*[@id="custtablelistpage_1_QuickFilterControl_listbox_item1"]'

    option = WebDriverWait(browser, 20).until(
        EC.visibility_of_element_located((By.XPATH, xpath))
        )

    browser.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});", option
        )

    ActionChains(browser).move_to_element(option).click().perform()

    

    WebDriverWait(browser, 15).until(
            EC.presence_of_element_located((By.XPATH, '(//*[@focusable="false"])[1]'))
            ).click()
    time.sleep(5)
    #open account details
    element = WebDriverWait(browser, 15).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="CustTable_AccountNum_3_0_0_input"]')
            )
        )

    ActionChains(browser).double_click(element).perform()

    print(f"Opening account details for {customer}...")
    time.sleep(15)

    max_attempts = 5

    for attempt in range(1, max_attempts + 1):
                
        for d in download_address:
            if os.path.exists(d):
                os.remove(d)
                print(f"Deleted file: {d}")
            else:
                print(f"File not found: {d}")
            
        number = random.randint(1, 10)
        numbere = random.randint(1, 10)
        start_date = (datetime.today() - timedelta(days=attempt+60+number)).strftime('%m/%d/%Y')
        end_date = (datetime.today() + timedelta(days=attempt+numbere)).strftime('%m/%d/%Y')
        try:
            print(f"Attempt {attempt} for {customer}...")

            # Click on the "Account Statement" button
            WebDriverWait(browser, 15).until(
                EC.presence_of_element_located((By.XPATH, '//*[@name="mibCustAccountStatementExt"]'))
                ).click()
            print(f"Generating account statement for {customer}...")

            # Select the date range for the statement
            sd= WebDriverWait(browser, 15).until(
                EC.presence_of_element_located((By.XPATH, '(//*[@aria-controls="ui-datepicker-div"])[1]'))
                )
            sd.click()
            sd.send_keys(Keys.CONTROL, "a")
            sd.send_keys(start_date)
            sd.send_keys(Keys.TAB)

            print(f"Selecting start date for {customer}...")
            time.sleep(2)
            # Select the end date for the statement
            ed= WebDriverWait(browser, 5).until(
                EC.presence_of_element_located((By.XPATH, '(//*[@aria-controls="ui-datepicker-div"])[2]'))
                )
            ed.click()
            ed.send_keys(Keys.CONTROL, "a")
            ed.send_keys(end_date)
            ed.send_keys(Keys.TAB)
            
            print(f"Selecting end date for {customer}...")
            time.sleep(2)

            # Click the "OK" button to generate the statement
            WebDriverWait(browser, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[contains(text(),"OK")]'))
                ).click()
            print(f"Account statement generation for {customer} has started.")
            time.sleep(5)  # Wait for the statement to be generated 
            #Click on the "Export" button to export the statement to Excel
            # Confirm processing operation
            WebDriverWait(browser, 15).until(
                 EC.presence_of_element_located(
                      (By.XPATH, '(//*[contains(text(),"Processing operation - Customer account statement")])[1]'))
                      )

            print(f"SUCCESS: Statement generation started for {customer}")
            break

        except TimeoutException:
            print(f"Attempt {attempt} unsuccessful for {customer}")

    WebDriverWait(browser, 215).until(
         EC.presence_of_element_located((By.XPATH, '//*[@class="button-label button-label-dropDown" and contains(@id,"PdfViewerExportMenuButton_label")]'))
         ).click()
    
    print(f"Clicking export button for {customer}...")

    #choose pdf option
    WebDriverWait(browser, 15).until(
         EC.presence_of_element_located((By.XPATH, '//*[@data-dyn-controlname="PdfViewerExportToPdfButton"]'))
         ).click()
    print(f"Exporting account statement to pdf for {customer}...")

    # WebDriverWait(browser, 15).until(
    #      EC.presence_of_element_located((By.XPATH, '//*[@data-dyn-controlname="SystemDefinedCloseButton"]'))
    #      ).click()
    

    # Wait for the file to download (this might need adjustments based on download time)
    WebDriverWait(browser, 5*60).until(
         lambda driver: len(glob.glob(f"{download_address}")) > len(f_path)
         )
    print("Export to PDF completed")


    # name the email subject
    subject = f"{customer} Accounts Statement - {formatted_date}"

    # Define the email function (dont call it email!)
    body = "Greetings,\nI trust this email finds you well.\nPlease find attached your statement for your reconciliation and subsequent payment upon completion.\nKindly ensure that this amount is settled promptly.\n\nYour cooperation will be greatly appreciated."
    # Define the file to attach

    # from YaleloProj.gmail_sender import gmail_function
    
    download_address = glob.glob(
        os.path.join(download_path, "Customer account statement*.pdf")
    )

    if download_address:
        attachment_path = download_address[-1]
        from gmail_sender import gmail_function
        gmail_function(
            f'{gmail_ac}, vguzman@yalelo.ug, btumwebaze@yalelo.ug, areasaleslead@yalelo.ug',
            subject_line=subject,
            email_body=body,
            attachment_path=attachment_path
        )
    else:
        print("No customer account statement PDF found.")
    for d in download_address:
        if os.path.exists(d):
            os.remove(d)
            print(f"Deleted file: {d}")
        else:
            print(f"File not found: {d}")
time.sleep(10)
kill_browser("chrome")