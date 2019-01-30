from selenium import webdriver
from selenium.webdriver.support.ui import Select
import xlrd 
loc = ("Hours.xlsx") 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
driver = webdriver.Chrome()
driver.get('#your_company_website')
user_name = driver.find_element_by_name('Skin$ctl08$LoginNameText')
user_name.send_keys('#your_username')
password = driver.find_element_by_name('Skin$ctl08$LoginPasswordText')
password.send_keys('#your_password')
login_button = driver.find_element_by_name('Skin$ctl08$ctl14')
login_button.click()
select_role = driver.find_elements_by_xpath("//*[contains(text(), 'Co-op')]")
select_role[1].click()
driver.find_element_by_xpath("//*[contains(text(), 'Go to time sheet')]").click()
for i in range(sheet.nrows):
    try:
        driver.find_element_by_xpath("//*[contains(text(), 'Add New Entry')]").click()
        select = Select(driver.find_element_by_name('Skin$body$ctl01$WDL'))
        select.select_by_index(int(sheet.cell_value(i, 0)))
        select = Select(driver.find_element_by_name('Skin$body$ctl01$StartDateTime1'))
        select.select_by_value(str(int(sheet.cell_value(i,1))).zfill(4))
        select = Select(driver.find_element_by_name('Skin$body$ctl01$EndDateTime1'))
        select.select_by_value(str(int(sheet.cell_value(i,2))).zfill(4))
        post_b = driver.find_element_by_name("Skin$body$ctl01$ctl08")
        post_b.click()
    except:
        driver.find_element_by_xpath("//input[@value='Cancel']").click()
        pass
		
driver.close()