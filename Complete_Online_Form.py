import openpyxl as op
from selenium import webdriver
from from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException

#Create a form in excel for users to complete. Add a button and assign a macro
#to it executing a shell command for this script.

xlsxLocation = '[EXCEL FILE PATH]'
wb = op.load_workbook(xlsxLocation, data_only=True, read_only=True)
PythonData = wb['PythonData']

DelPy = "#"

website = ''

error_log = []

def data_pull(pull_range):
     data_list = []
     for cell in pull_range: #specify pull range variables to avoid errors
          for x in cell:
               if x.value != 'NO':
                   data_list.append(x.value)
     return data_list

 def Add_Note():
    driver.find_element(By.PARTIAL_LINK_TEXT, 'Add').click()
    driver.switch_to.alert.accept()
    
def send_Xp(Xp, accountInformation):
    driver.find_element(By.XPATH, Xp).send_keys(accountInformation)
    
def send_name(name, accountInformation)
    driver.find_element(By.NAME, name).send_keys(accountInformation)

def click_xpath(Xp):
    driver.find_element(By.XPATH, Xp).click()

def click_name(name)
    driver.find_element(By.NAME, name).click()

def dropdown_select(Xp, dropdown_selection):
    site_dropdown = driver.find_element(By.XPATH, Xp)
    site_select = Select(site_dropdown)
    site_select.select_by_visible_text(dropdown_selection)

UserID = PythonData['B3'].value
UserPassword = PythonData['B4'].value
FileType = PythonData['B5'].value

excelRange_accounts = PythonData['B6' : 'B100']
excelRange_notes = PythonData['b101' : 'B200']

accounts = data_pull(excelRange_accounts)
notes = data_pull(excelRange_notes)

driver = webdriver.Chrome()
driver.implicitly_wait(10)
driver.get(website)

#Login
send_name(r'user', UserID)
send_name(r'PASSWORD', UserPassword)
click_name(r'submit')
click_name(r'button')
click_xpath(r'/html/body/form/center/input')

click(r'/html/body/div[2]/table/tbody/tr[2]/td/center/table/tbody/tr[2]/td/ul/li/a/font/b')

for x in accounts:
    try:
        if x != "NO":
            send_name(x.split(DelPy, 2)[0])
            send_name(x.split(DelPy, 2)[1])
    else:
        error_log.append(x.split(DelPy, 2)[2])

click(r'/html/body/div[2]/table/tbody/tr[2]/td/center/table/tbody/tr[3]/td/ul/li/a/font/b')

for x in notes:
    Add_Note()
    send_Xp(x.split(DelPy, 1)[0])
    dropdown_select(x.split(DelPy, 1)[1])

if len(error_log) > 0:
    print("Errors ocurred with the following accounts:")
    for x in error_log:
        print(x)
