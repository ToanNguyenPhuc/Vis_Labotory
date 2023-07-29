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
from GlobalLink import GlobalLinks


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
        self.links = GlobalLinks()
        self.CountDone = 0
        self.CountCrawlTableError = 0
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
        with open(output,'a',newline='',encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(columnsName)            
    def getTableContent(self):
        ''' Crawl table value from website

            Parameter
            ---------
             
            ---------
            Returns
            -------
            df[ColumnsName] : list

         '''
        tbodyContainContent = self.driver.find_element(By.XPATH,'//*[@id="pageHistoricQuotes"]/div[3]/div[1]/table/tbody')
        rowsList = tbodyContainContent.find_elements(By.TAG_NAME,'tr')
        for row in rowsList[1:]:
            list_ = [None] * 7
            tdList = row.find_elements(By.TAG_NAME,'td')
            tdTextList = [tdEle.text for tdEle in tdList]
            self.saveData(tdTextList,self.output_csv)
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
        ListISIN = self.getISIN(self.links.input)
        for ISIN in ListISIN:
            self.output_csv = f'ArivaXetraDelisting/{ISIN}.csv'
            self.output_excel = f'ArivaXetraDelisting/{ISIN}.xlsx'
            self.writeColumnName(self.output_csv)
            url = f'https://www.ariva.de/{ISIN}/kurse/historische-kurse?go=1&boerse_id=6&month=2023-07-31&currency=EUR'
            self.getWebsite(url)
            time.sleep(0.2)
            try: 
                self.getTableContent()
                self.convertToExcel(self.output_csv,self.output_excel)
                self.CountDone+=1
            except Exception as e:     
                # Log the error message using the 'error' level
                logging.error(f"Error occurred at {ISIN}: {e}")
                self.CountCrawlTableError+=1
            print(f'Current Company: {ISIN} _ TotalDone: {self.CountDone} _ TotalError: {self.CountCrawlTableError} _ Total: {self.CountDone + self.CountCrawlTableError}')
crawler = Crawler()
crawler.MainHandle()

