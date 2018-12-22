#Importing the necessary packages/modules to perform the script
import time
import xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


print("This is a python script that grabs queried Historical Stock Values from www.YahooFinance.com, " 
      + "prints it out and puts it into an excel file and names the file according to the Stock Name")

#Create our first Chrome webdriver and go to the Yahoo Finance website
url = 'https://finance.yahoo.com/'
options = Options()
options.add_argument('--headless')
driver = webdriver.Chrome(options = options)
driver.get(url)
html = driver.page_source

#Get the stock name and find the respective, Yahoo Finance historical data URL
stock_name = input("Please enter the stock ticker symbol: ")
action = ActionChains(driver)
Search_Box = driver.find_element_by_xpath("//input[@placeholder='Quote Lookup']")
action.move_to_element(Search_Box)
action.click(Search_Box)
action.send_keys(stock_name)
action.send_keys(Keys.ENTER)
action.perform()
element = WebDriverWait(driver, 20).until(
EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Historical Data')]")))
element.click()
time.sleep(1)

#Create a new webdriver with the queried stock's Historical Data URL
new_html=driver.current_url
driver = webdriver.Chrome()
driver.get(new_html)
body_tag = driver.find_element_by_tag_name("body")
time.sleep(1)

#Page down via selenium so the page can load all the html data via JavaScript
no_of_pagedowns = 15
while no_of_pagedowns:
	body_tag.send_keys(Keys.PAGE_DOWN)
	time.sleep(0.1)
	no_of_pagedowns-=1
time.sleep(1)

#Create our html parser and find the data
html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
allData = soup.find_all('tr',class_='BdT Bdc($c-fuji-grey-c) Ta(end) Fz(s) Whs(nw)')

#Create our Excel notebook.
my_xls=xlwt.Workbook()
my_sheet=my_xls.add_sheet(stock_name)
row_num=0

#Create our Column Headers
my_sheet.write(row_num,0,"Date")
my_sheet.write(row_num,1,"Opening Value")
my_sheet.write(row_num,2,"High Value")
my_sheet.write(row_num,3,"Low Value")
my_sheet.write(row_num,4,"Close Value")
my_sheet.write(row_num,5,"Adj Close Value")
my_sheet.write(row_num,6,"Volume")
row_num+=1

#Create a loop which gathers all the necessary text via our parser
for data in allData:
    date = data.td.text #gathers the date via 
    stock_row = data.find_all('td', class_='Py(10px) Pstart(10px)')
    
    try:
        stock_values = stock_row[0].text
        stock_values_1 = stock_row[1].text
        stock_values_2 = stock_row[2].text
        stock_values_3 = stock_row[3].text
        stock_values_4 = stock_row[4].text
        stock_values_5 = stock_row[5].text

#Write the Data to our Excel File
        print(date + " " + stock_values + " " + stock_values_1 + " " 
              + stock_values_2 + " " + stock_values_3 + " " + stock_values_4 + " " + stock_values_5)
        my_sheet.write(row_num,0,date)
        my_sheet.write(row_num,1,stock_values)
        my_sheet.write(row_num,2,stock_values_1)
        my_sheet.write(row_num,3,stock_values_2)
        my_sheet.write(row_num,4,stock_values_3)
        my_sheet.write(row_num,5,stock_values_4)
        my_sheet.write(row_num,6,stock_values_5)
        row_num+=1
    except:
        row_num+=1

#Save the Data to an Excel File and name the File corresponding to the stock name
my_xls.save(stock_name + ".xls")
print("The data was saved into the file " + stock_name 
      + ".xls, in this program's working directory")