#-----------------------Libraries---------------------------------------------
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import numpy as np
import pandas as pd
#-----------------------Perm variables----------------------------------------
DelPy = "#" #python delimiter

Website = """
very long URL that would
take multiple 
lines.
"""

error_log = []
data_listing = []

#Data storage, as needed
DataFile = r''
backup_file = r'' #specify file path
output_file = r'' #specify file path

#read data into memory
xlsx = pd.read_excel(DataFile, sheet_name="PyData")
df = xlsx['ColumnA'] + DelPy + xlsx['ColumnB'].astype(str)

#-----------------------Functions---------------------------------------------
def send(Xp, Information):
    driver.find_element(By.XPATH, Xp).send_keys(Information)

def click(Xp):
        driver.find_element(By.XPATH, Xp).click()

def click_link(exact_text_of_link):
     driver.find_element(By.LINK_TEXT, exact_text_of_link).click()

#-----------------------Begin Execution---------------------------------------

# Set up web driver object
driver = webdriver.Chrome()
driver.implicitly_wait(10)
original_window = driver.current_window_handle

# Iterate through the list to pull data from a dropdown assigned to users
for x in df:
    try:
        driver.switch_to.window(original_window)
        driver.switch_to.new_window("window")
        # Navigate to the website
        driver.get(Website) 
        # Login
        driver.find_element(By.NAME, r'user').send_keys(x.split(DelPy, 1)[0])
        driver.find_element(By.NAME, r'PASSWORD').send_keys(x.split(DelPy, 1)[1])
        driver.find_element(By.NAME, r'submit').click()
        driver.find_element(By.NAME, r'button').click()
        driver.find_element(By.PARTIAL_LINK_TEXT, r'MAIN').click()
        driver.find_element(By.XPATH, r'/html/body/form/center/input').click()
        #web elements to interact with on the site
        new_link = r'/html/body/form/div/table/tbody/tr[2]/td/input'
        continue_button = r'/html/body/form/div/table/tbody/tr[2]/td/center/table/tbody/tr[3]/td/input'
        dropdown = r'/html/body/form/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/select'
        click_link(new_link)
        click(continue_button)
        dropdown = driver.find_element(By. XPATH, dropdown)
        options = dropdown.find_elements(By.TAG_NAME, "option")
        for option in options:
                if option.get_attribute("value") != "none": #skip the first option
                    varPull = x.split(DelPy, 1)[0] + DelPy + option.get_attribute("value")
                    data_listing.append(varPull.strip())
    except NoSuchElementException:
        print("Encountered an error")
        error_log.append(x.split(DelPy, 1)[0])
    finally:
        # Close the browser window
        driver.close()

#close the driver object
driver.switch_to.window(original_window)
driver.quit()

#write dropdown data to excel
df2 = pd.DataFrame(data_listing, columns=['original'])
df2.to_excel(output_file, sheet_name="RawData", index=False)
df2.to_excel(backup_file, sheet_name="RawData", index=False)

#print the errors
if len(error_log) > 0:
     print("errors for the following users:")
     for x in error_log:
          print(x)
