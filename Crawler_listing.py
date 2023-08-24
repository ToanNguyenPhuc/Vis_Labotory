# This file using ISIN to crawl price of company from Ariva webpage
from selenium import webdriver
# selenium-wire
from selenium.webdriver.chrome.service import Service as ChromeService
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import re
import csv
import time
from selenium.webdriver.chrome.options import Options

# selenium 4
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import requests
import logging
import openpyxl


# Configure logging settings (optional)
logging.basicConfig(
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filename="error_log.txt",
    filemode="a"
)


class Crawler:
    def __init__(self) -> None:
        self.chrome_options = Options()
        self.chrome_options.add_argument("--headless")
        self.driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        self.CountDone = 0
        self.input = 'Xetra_listing.csv'
        self.fail_Crawl_Report_File_Path = 'ERROR_CRAWL_DIVIDEND_LISTING'
        self.CountCrawlTableError = 0
        self.CountSuccess = 0
        self.CountFail = 0
        self.start_PATH = 'https://www.asx.com.au/markets/trade-our-cash-market/directory'
        pass
    def getWebsite(self,url):
        self.driver.get(url)
    def getISIN(self,FilePath: str ,ColumnsName = 'ISIN'):
        ''' Get table from csv file then return columns contain ISIN

            Parameter
            ---------
                FilePath (str) : Path to CSV file
                ColumnsName (str) : Name of columns contain ISIN (defauls : 'ISIN')
            ---------
            Returns
            -------
            df[ColumnsName] : obj

         '''
        if 'csv' in FilePath:
            df = pd.read_csv(FilePath)
            return df[ColumnsName]
        elif 'xlsx' in FilePath:
            df = pd.read_excel(FilePath)
            return df[ColumnsName]
    def writeColumnName(self,output):
        columnsName = ['Datum','Erster','Hoch','Tief','Schluss','Stücke','Volumen']
        with open(output,'w',newline='',encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(columnsName)            
    def getTableContent(self,output):
        ''' Crawl table value from website

            Parameter
            ---------
             
            ---------
            Returns
            -------
            df[ColumnsName] : list

         '''
        table_ele = self.driver.find_element(By.XPATH,'//*[@id="tablesorter_browse"]/tbody')
        row_eles = table_ele.find_elements(By.CSS_SELECTOR,'td')
        for row in row_eles:
            print(row.text)
    def saveData(self,data,output):
        ''' Save data value to file

            Parameter
            ---------
                data (list): data to save
                output (string): path name of the file wanna save
            ---------
            Returns
            -------
            None
            
         '''
        with open(output,'a',newline='',encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(data)
    def getReport(self,state,company_symbol):
        ''' Crawl table value from website

            Parameter
            ---------
            state: State of cawler with current company
            company_symbol: Symbol of current company 
            ---------
            Returns 
            -------
            

        '''
        if state == 'Success':
            self.CountSuccess +=1 
        elif state == 'Fail':
            self.CountFail +=1
        print(str(company_symbol).center(20,'-'))
        print('state:'+str(state).rjust(14,'.'))
        print('Success:'+str(self.CountSuccess).rjust(12,'.'))
        print('Fail:'+str(self.CountFail).rjust(15,'.'))
    def handleFailCrawl(self,symbol,e):
        with open(self.fail_Crawl_Report_File_Path,'a',newline='',encoding= 'utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow([str(symbol).strip()])
        logging.error(f"Error occurred at {symbol}: {e}")
    def convertToExcel(self,input,output):
        ''' Convert csv file to excel file

        Parameter
        ---------
            input : path of csv file
            output : path of excel file
        ---------
        Returns
        -------
        None
        
        '''
        data = pd.read_csv(input)
        df = pd.DataFrame(data=data)
        df.to_excel(output)
    def MainHandle(self):
        self.getWebsite(self.start_PATH)
        time.sleep(1)
        cookie_ele = self.driver.find_element(By.XPATH,'//*[@id="onetrust-accept-btn-handler"]').click()
        button_ele = self.driver.find_element(By.XPATH,'//*[@id="company_directory"]/div/div[1]/div[2]/div[2]/a/span/span[2]').click()
        input('')
    
crawler = Crawler()
crawler.MainHandle()

