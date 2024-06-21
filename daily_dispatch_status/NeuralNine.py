#%%#import modules
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

#%%

Options=Options()
Options.add_experimental_option("detach",True)

driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=Options )
# driver.get("https://app.powerbi.com/")
driver.get("https://www.neuralnine.com/")
#maximise window
driver.maximize_window()

#%%
#get link of all books
links=driver.find_elements("xpath", "//a[@href]")
print(links)
for link in links:
    if "Books" in link.get_attribute("innerHTML"):
        link.click()
        break
book_links=driver.find_elements("xpath",
    
                                "//div[contains(@class,'elementor-column-wrap')][.//h2[text()[contains(., '7 IN 1')]]]")

#click the link
book_links[0].click()

# #switch window
# driver.switch_to.window(driver.window_handles[1])
# time.sleep(3)
# #%%
# #get price details
# buttons=driver.find_elements("xpath", 
#                              "//a[.//span[text()[contains(.,'Taschenbuch')]]]//span[text()[contains(., '£')]]")

# for button in buttons:
#     print(button.get_attribute("innerHTML").replace("&nbsp;"," "))

#%%
time.sleep(5)
driver.quit()