from selenium.webdriver.common.by import By

import XLUtils2
from selenium import webdriver
import time


driver = webdriver.Chrome(executable_path="C:\BrowserDrivers\chromedriver.exe")
driver.get("http://computer-database.herokuapp.com/computers")

path="C:/TestData/Add-A_New_Computer1.xlsx"


rows=XLUtils2.getRowCount(path,"Sheet1")
print("the row count is ", rows)

for r in range(2,rows+1):
    Computer_name = XLUtils2.readData(path, "Sheet1", r, 1)
    Introduced_date = XLUtils2.readData(path, "Sheet1", r, 2)
    Discontinued_date= XLUtils2.readData(path, "Sheet1", r, 3)
    Company = XLUtils2.readData(path, "Sheet1", r, 4)

    driver.find_element(By.ID, "add").click()
    Computer_name=driver.find_element(By.ID, "name").send_keys(Computer_name)
    Introduced_date=driver.find_element_by_name("introduced").send_keys(Introduced_date)
    Discontinued_date=driver.find_element(By.ID, "discontinued").send_keys(Discontinued_date)
    Company=driver.find_element_by_id("company").send_keys(Company)

    driver.find_element(By.CSS_SELECTOR, ".primary").click()
    time.sleep(3)


    Status = driver.find_element(By.CSS_SELECTOR, "#main > h1").text
    print(Status)


    if driver.title == "Computers database":
        driver.save_screenshot("C:/TestData/TestResults/PageTitle_Passed.png")
        XLUtils2.writeData(path, "Sheet1", r, 5, "Test Failed")
    else:
        driver.save_screenshot("C:/TestData/TestResults/PageTitle_Failed.png")
        XLUtils2.writeData(path, "Sheet1", r, 5, "Test Passed")

    driver.quit()


