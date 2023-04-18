from selenium.webdriver.support.select import Select
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import os 
import time
import pandas as pd


driver = webdriver.Chrome(r"C:\Users\mtl91475\Desktop\Coding\pySelenium\chromedriver.exe")
driver.get("http://163.50.57.101/FC008/FractionCombineMasterLotHist.aspx")

inputElement = driver.find_element(By.ID, 'MainContent_txtNewLot')
inputElement.send_keys('1')