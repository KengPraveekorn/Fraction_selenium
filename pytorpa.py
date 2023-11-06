import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup 
import openpyxl as OP

# driver = webdriver.Edge(r"C:\Users\mtl91475\Desktop\Coding\pySelenium\msedgedriver.exe")
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()))
driver.get("http://mtl91475:Gangkia@11@163.50.57.101/FC005/S000103.aspx")

select = Select(driver.find_element(By.ID, 'MainContent_lbPTM0014')) # Process
select2 = Select(driver.find_element(By.ID, 'MainContent_lbPTC0006')) # Input Code

select.deselect_by_value("All")
select2.deselect_by_value("All")


# select.select_by_value("2191") # SMT
select.select_by_value("2180") # Outgoing Inspection
select2.select_by_value("12") # Fraction Combine

driver.find_element(By.ID, 'MainContent_btnRefresh').click()

##------------------------------------------------------------------------------------------------------------##
############################################# Get from Wip ####################################################

data = driver.page_source
dthtml = pd.read_html(data)[3]
df = pd.DataFrame(dthtml)

##------------------------------------------------------------------------------------------------------------##
############################################# Get from Fraction and write to excel ####################################################

dff = df[1].drop(0)
dfl = len(dff)
dff.to_csv("LotWip.csv")

i=0
while i < dfl:
    lotno = df[1][i+1]
    driver.get("http://mtl91475:Gangkia@11@163.50.57.101/FC008/FractionCombineMasterLotHist.aspx")
    driver.find_element(By.ID, 'MainContent_txtNewLot').send_keys(lotno)
    driver.find_element(By.NAME, 'ctl00$MainContent$btnRefresh').click()
    soup = BeautifulSoup(driver.page_source)
    soup_table = soup.find_all("table")
    dttable = pd.read_html(str(soup_table))[1]
    dtfrac = dttable["FRACTIONLOT"][0]
    print(dtfrac)
    excel_file = (r"F:\MT900\900 Public\MT920\002_Outgoing inspection\For_RPA_00178\Lotfrac.xlsx")
    wb = OP.load_workbook(excel_file)
    sheet = wb.active
    sheet.delete_cols(i+1)
    sheet.cell(row=i+1,column=1).value = dtfrac
    wb.save(excel_file)
    i += 1



