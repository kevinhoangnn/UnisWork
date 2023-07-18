import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from difflib import SequenceMatcher
from pandas import read_excel
import os.path
import json
from glob import glob
from BNPauto import facilityMatcher

def get_bnp_name(account):
    invoice_accs = read_excel('BNP Excel Sheet.xlsx', sheet_name='Account_Fac_Freq')
    acc_name = invoice_accs[invoice_accs['AccountName'] == account]
    acc_name = acc_name['BNP Account Name'][0]

    return acc_name

def csr_update_item(invoice_num):
    with open('accountconfigs.json', 'r') as f:
        data = json.loads(f.read())

    chromeOptions = webdriver.ChromeOptions()
    prefs = {'safebrowsing.disable_download_protection' : True}
    chromeOptions.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chromeOptions)
    action = ActionChains(driver)
    driver.get(data['bnpTestDomain'])

    driver.maximize_window()

    #Logging into BNP
    interactor = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'inputUserName')))
    interactor.send_keys(data['bnpUser'])
    interactor = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'inputPassword')))
    interactor.send_keys(data['bnpPass'])
    interactor = driver.find_element(By.XPATH,"/html/body/div/footer/div/button")
    interactor.click()

    #Clicking Sales Module from Module Dropdown Menu
    select = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'TS_span_menu')))
    select.click()
    select = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="headmenu_mn_active"]/div/ul/li[1]')))
    select.click()

    #Clicking Invoice Management from Invoice Dropdown Menu
    select = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="headmenu"]/li[3]/span')))
    select.click()
    select = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/header/div[1]/ul/li[3]/div/ul/li[1]/a')))
    select.click()

    #Inputting information on the Invoice Management Page and Searching

    #Invoice Number 
    interactor = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[1]/div/input')
    interactor.send_keys(invoice_num)

    #Clicking Search
    interactor = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[3]/div[11]/button')
    interactor.click()
    interactor = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[3]/div[11]/button')
    interactor.click()

    detail = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, 'Detail')))
    detail.click()

    time.sleep(10)

if __name__ == '__main__':
    csr_update_item(19029651)