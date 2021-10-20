from selenium.webdriver.common.by import By

import XLUtils2
from selenium import webdriver
import time


driver = webdriver.Chrome(executable_path="C:\BrowserDrivers\chromedriver.exe")
driver.get("http://computer-database.herokuapp.com/computers")

path="C:/TestData/Computer_Database.xlsx"

rows=XLUtils2.getRowCount(path,"Sheet1")
print("the row count is ", rows)

for r in range(2,rows+1):
    Computer_Database = XLUtils2.readData(path, "Sheet1", r, 1)

    driver.find_element_by_id("searchbox").clear()
    userName = driver.find_element_by_id("searchbox").send_keys(Computer_Database )
    driver.find_element_by_id("searchsubmit").click()
    time.sleep(3)

    Status = driver.find_element(By.CSS_SELECTOR, "#main > h1").text
    print(Status)

    if driver.title == "Computers database":
        driver.save_screenshot("C:/TestData/TestResults/PageTitle_Passed.png")
        XLUtils2.writeData(path, "Sheet1", r, 2, "Test Passed")
    else:
        driver.save_screenshot("C:/TestData/TestResults/PageTitle_Failed.png")
        XLUtils2.writeData(path, "Sheet1", r, 2, "Test failed")

# Validating cancel operation
    driver.find_element(By.CSS_SELECTOR, "tr:nth-child(1) > td > a").click()
    driver.find_element(By.CSS_SELECTOR, ".btn:nth-child(2)").click()

   # driver.find_element_by_css_selector("#Layer_1 > g > path:nth-child(5)").click()
#  driver.quit()
