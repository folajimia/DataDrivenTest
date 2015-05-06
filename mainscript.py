from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from mmap import mmap,ACCESS_READ
from xlrd import open_workbook,cellname
from decimal import *
import unittest, time, re



def ExcelOpener(FileLocation):
    priceList=[]
    book=open_workbook(FileLocation)
    sheet=book.sheet_by_index(0)
    for j in range(1,sheet.nrows):
        if sheet.cell_value(j,0)=="Y".upper():
            print (sheet.cell_value(j,4))
            if sheet.cell_value(j,4)=="OpenBrowser":
                driver=OpenBrowser(sheet.cell_value(j,7))
            elif sheet.cell_value(j,4)=="navigate_to":
                print (driver,sheet.cell_value(j,7))
                navigate_to(driver,sheet.cell_value(j,7))
            elif sheet.cell_value(j,4)=="click_element":
                click_element(driver,sheet.cell_value(j,5),sheet.cell_value(j,6))
            elif sheet.cell_value(j,4)=="send_keys":
                send_keys(driver,sheet.cell_value(j,5),sheet.cell_value(j,6),sheet.cell_value(j,7))
            elif sheet.cell_value(j,4)=="verify_element":
                verify_element(driver,sheet.cell_value(j,5),sheet.cell_value(j,6))
            elif sheet.cell_value(j,4)=="store_text":
                price = store_text(sheet.cell_value(j,5),sheet.cell_value(j,6))
                pricex=price.replace('$','').replace(',','')
                priceF=float(pricex)
            else:
                print ("it ends here")
            close_browser(driver)

def OpenBrowser(browserType):
    if browserType.lower()=='firefox':
        driver = webdriver.Firefox()
    elif browserType.lower()=='ie':
        driver = webdriver.Ie()
    elif browserType.lower()=='chrome':
        driver = webdriver.Chrome()
    elif browserType.lower()=='opera':
        driver = webdriver.Opera()
    else:
        print ("Browser is not available")
    return driver

def navigate_to(driver, url):
    driver.get(url)
    return

def send_keys(driver, locator, locString, data):
    if locator=="xpath":
        driver.find_element_by_xpath(locString).clear()
        driver.find_element_by_xpath(locString).send_keys(data)
    elif locator=="name":
        driver.find_element_by_name(locString).clear()
        driver.find_element_by_name(locString).send_keys(data)
    elif locator=="id":
        driver.find_element_by_id(locString).clear()
        driver.find_element_by_id(locString).send_keys(data)
    else:
        print ("locator is not available")
    return


def click_element(driver, locator, locString):
    if locator=="xpath":
        driver.find_element_by_xpath(locString).click()
    elif locator=="name":
        driver.find_element_by_name(locString).click()
    elif locator=="id":
        driver.find_element_by_id(locString).click()
    else:
        print ("locator is not available")
    return

def verify_element(driver, locator, locString):
    if locator=="xpath":
        wait= WebDriverWait(driver,90)
        wait.until(EC.presence_of_element_located((By.XPATH,locString)))
    elif locator=="name":
        wait= WebDriverWait(driver,90)
        wait.until(EC.presence_of_element_located((By.NAME,locString)))
    elif locator=="id":
        wait= WebDriverWait(driver,90)
        wait.until(EC.presence_of_element_located((By.ID,locString)))
    else:
        print ("locator is not available")
    return

def store_text(driver, locator, locString):
    if locator=="xpath":
        price = driver.find_element_by_xpath(locString).text
    elif locator=="name":
        price = driver.find_element_by_name(locString).text
    elif locator=="id":
        price = driver.find_element_by_id(locString).text
    else:
        print ("locator is not available")
    return price

def close_browser(driver):
    driver.quit()
    return


ExcelOpener('C:\\Downloads\\keyworddrivengug.xls')
