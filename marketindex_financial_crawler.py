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
import undetected_chromedriver as uc
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
        self.input = 'List_com_ASX.csv'
        self.fail_Crawl_Report_File_Path = 'ERROR_CRAWL_financial_maarketindex'
        self.CountCrawlTableError = 0
        self.CountSuccess = 0
        self.CountFail = 0
        self.start_PATH = 'https://www.marketindex.com.au/asx/symbol/financials?print'
        self.output = 'financial_maarketindex.csv'
        pass
    def getWebsite(self,url):
        self.driver.get(url)
    def getCompanySymbol(self,FilePath: str ,ColumnsName = 'ISIN'):
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
        columnsName = ['Name','Shares Issued','Sector','Date Listed']
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
        table_elements = self.driver.find_elements(By.CSS_SELECTOR,'table')
        for ii,table_ele in enumerate(table_elements):
                print(f'table :{ii}')
                thead_elements = table_ele.find_element(By.CSS_SELECTOR,'thead')
                th_elements = thead_elements.find_elements(By.CSS_SELECTOR,'th') 
                index_value = [th_ele.text for th_ele in th_elements]
                self.saveData(index_value,output)

                # Handle table content
                tbody_elements = table_ele.find_element(By.CSS_SELECTOR,'tbody')
                tr_elements = tbody_elements.find_elements(By.CSS_SELECTOR,'tr')
                for tr_ele in tr_elements:
                        td_elements = tr_ele.find_elements(By.CSS_SELECTOR,'td')
                        td_elements_value = [td_ele.text for td_ele in td_elements]
                        self.saveData(td_elements_value,output)
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
    def logIn(self,url):
                # driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        tk = 'toankutehip@gmail.com'
        mk = 'vislab123'
        self.driver = uc.Chrome()
        self.driver.get(url)
        email_input = self.driver.find_element(By.XPATH,'//*[@id="authentication-root"]/div/div[2]/form/input[2]')
        email_input.send_keys(tk)
        password_input = self.driver.find_element(By.XPATH,'//*[@id="authentication-root"]/div/div[2]/form/input[3]')
        password_input.send_keys(mk)
        logIn_btn = self.driver.find_element(By.XPATH,'//*[@id="authentication-root"]/div/div[2]/form/input[4]')
        logIn_btn.click()
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
    def getFullFinancial(self):
        print_financial_btn = self.driver.find_element(By.XPATH,'//*[@id="summary-pane"]/div[1]/div/a')
        print_financial_btn.click()
    def MainHandle(self):
        self.writeColumnName(self.output)
        self.logIn('https://www.marketindex.com.au/login')
        listing_symbols = self.getCompanySymbol('List_com_ASX.csv','Listing')
        for symbol in listing_symbols:
            self.output = f'financial/{symbol}.csv'
            if str(symbol) != 'nan' :
                url = self.start_PATH.replace('symbol',str(symbol).lower())
                self.getWebsite(url)
                time.sleep(1)
                try:
                    self.getTableContent(self.output)
                    self.getReport('Success',symbol)
                except Exception as e:
                    self.handleFailCrawl(symbol,e)
                    self.getReport('Fail',symbol)
            # self.getTableContent()
crawler = Crawler()
crawler.MainHandle()
