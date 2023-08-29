"""
Ban Crawl nay phai su dung symbol cua cong ty nhu la acwn / khong the dung ISIN
"""
from selenium import webdriver
# selenium-wire
from selenium.webdriver.chrome.service import Service as ChromeService
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.common.action_chains import ActionChains
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
# from GlobalLink import GlobalLinks


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
        self.columns_name = ['Kürzel',
                'Sector',
                'Industry',
                'Investment Style',
                'Shares Outstanding']
        self.File_path_to_company_symbol_file_contain = 'listing_infor.csv'
        self.company_symbol_column_name = 'ISIN'
        self.CountSuccess = 0
        self.CountFail = 0
        self.output = 'morning_star_Xetra_listing'
        self.error_company_symbols_list = 'error_symbols.csv'
        pass
    def getWebsite(self,url):
        self.driver.get(url)
    def getCompanySymbol(self,FilePath: str ,ColumnsName = 'ISIN'):
        ''' Get table from csv file then return columns contain Company Symbol

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
        with open(output, 'a', encoding= 'utf-8' , newline = '') as file:
            writer = csv.writer(file)
            writer.writerow(self.columns_name)
    def getTableContent(self,symbol):
        ''' Crawl table value from website

            Parameter
            ---------
             
            ---------
            Returns
            -------
            df[ColumnsName] : list

         '''
        dict_to_save = {'Symbol': symbol}
        UL_container = self.driver.find_element(By.XPATH,'//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div/div/div[1]/section[1]/div/div[2]/div[2]/ul')
        li_elements = UL_container.find_elements(By.TAG_NAME, 'li')
        for li in li_elements:
            column_name = li.text.split('\n')[0]
            value = li.text.split('\n')[1]
            if column_name in self.columns_name:
                dict_to_save[column_name] = value
        btn_ele = self.driver.find_element(By.XPATH,'//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div/div/div[1]/section[1]/div/div[2]/div[1]/button[3]')
        btn_ele.click()
        Shares_Outstanding_value = self.driver.find_element(By.XPATH,'//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div/div/div[1]/section[1]/div/div[2]/div[2]/ul/li[7]/span[2]')
        dict_to_save['Shares Outstanding'] = Shares_Outstanding_value.text
        self.saveData(dict_to_save.values(),self.output)

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
    def handleFailCrawl(self,symbol,e):
        with open(self.error_company_symbols_list,'a',newline='',encoding= 'utf-8') as file:
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
        Symbols = self.getCompanySymbol(self.File_path_to_company_symbol_file_contain,self.company_symbol_column_name)
        for symbol in Symbols:
            PATH = f'https://www.morningstar.com/stocks/xetr/{str(symbol).strip()}/quote'
            self.driver.get(PATH)
            time.sleep(3)
            try:
                self.getTableContent(str(symbol).strip())
                self.getReport('Success',symbol)
            except Exception as e:
                self.handleFailCrawl(symbol,e)
                self.getReport('Fail',symbol)
            
crawler = Crawler()
crawler.MainHandle()

