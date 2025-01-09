from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import *

from selenium.webdriver.chrome.service import Service

import time
import os
import sys
import random
import string
import re
import glob
import shutil
import csv

import pandas as pd
import numpy as np
import chromedriver_autoinstaller

from string import *
from random import randint
from os import path
from datetime import date, timedelta, datetime
#import xlrd

from openpyxl import load_workbook
import openpyxl
from openpyxl import Workbook

#chrome_options = Options()
#chrome_options.set_headless()
#driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
#driver = webdriver.Chrome(service=webdriver.chrome.service.Service(executable_path="C:\webdrivers\chromedriver.exe"))

service = Service(ChromeDriverManager().install())

#service = Service('C:\webdrivers\chromedriver.exe')

#driver = webdriver.Chrome('C:\webdrivers\chromedriver.exe')
driver = webdriver.Chrome(service = service)

driver.get("xxx")
time.sleep(3)
driver.refresh()
time.sleep(2)
driver.refresh()
driver.get("xxx")

driver.maximize_window()

wait = WebDriverWait(driver, 30)

sscorp = wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[@data-file='SUPPLYCORP']")))
sscorp.click()

hub = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[@data-file='02_HUB']")))
hub.click()

d_load = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[@data-file='DOWNLOAD']")))
d_load.click()

etl_log = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[@data-file='ARC_ETL_Log']")))
etl_log.click()
time.sleep(15)

search = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 'searchbox']")))
search.click()

#action = ActionChains()
#action.send_keys(search)
passerelle_download = 'C:/_____Passerelle_____'+'/'
if not os.path.exists(passerelle_download):
    os.makedirs(passerelle_download)
elif os.path.exists(passerelle_download):
    pass

df = pd.read_excel('C:/_____Passerelle_____/Qliksense.xlsx', header = 0)

filename = df['File Name'].tolist()
print()
print(len(filename))

def create_fol(country):
    global newpath
    #newpath = passerelle_download + str(country) + '/' + '_' + str(datetime.now().strftime("%a"))
    newpath = passerelle_download + str(country) + '/'

    if not os.path.exists(newpath):
        os.makedirs(newpath)

    elif os.path.exists(newpath):
        for filename in os.listdir(newpath):
            if os.path.isfile(os.path.join(newpath, filename)):
                os.remove(os.path.join(newpath, filename))

create_fol('Slovakia')
create_fol('Spain')
create_fol('Poland')
create_fol('Belgium')
create_fol('Romania')
create_fol('France')

for i in range(len(filename)):
    FN = filename[i][:-4]
    print(filename[i][:-4], "Count ---> ", i)
    search.send_keys(filename[i][:-4])
    time.sleep(3)
    try:
        search.send_keys(Keys.ENTER)
    except StaleElementReferenceException as e:
        search.send_keys(Keys.ENTER)
        
    fichier = wait.until(EC.element_to_be_clickable((By.XPATH, "//tbody[@id = 'fileList']")))
    try:
        fichier.click()
    except TimeoutException as e:
        driver.refresh()
        time.sleep(5)
        search.send_keys(filename[i][:-4])
        time.sleep(3)
        search.send_keys(Keys.ENTER)
        fichier.click()
        
    except ElementNotInteractableException as e:
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[text() = 'SUPPLYCORP/02_HUB/DOWNLOAD/ARC_ETL_Log/']"))).click()

    print(filename[i][11:13])
    time.sleep(2)
    country = filename[i][11:13]

    search = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id = 'searchbox']")))
    search.clear()
    time.sleep(1)
time.sleep(3)

def partition_files():
    files = glob.glob('C:/Users/'+ str(os.getlogin()) +'/Downloads/SPIC_KO_*.log')
    for file in files:
        file_name = 'C:/Users/'+ str(os.getlogin()) +'/Downloads/' + str(filename[i][0:33]) + '.log'
        print(file_name)

        def convertir_xl(pays):
            with open(file, encoding="utf-8") as handle:
                content = handle.readlines()
                info = content[14:]
            with open('C:/_____Passerelle_____/' + str(pays) + '/' + str(file[28:-3]) + str('csv'), 'w') as x:
                for line in info:
                    x.write(line)

        def move_to_folder(pays):
            xl = pd.read_csv(r'C:/_____Passerelle_____/' + str(pays) + '/' + str(file[28:-3]) + str('csv'), delimiter = ';', encoding="latin-1", engine='python')
            xl = pd.read_csv(r'C:/_____Passerelle_____/' + str(pays) + '/' + str(file[28:-3]) + str('csv'), delimiter = '|', encoding="latin-1", engine='python')    

        if file[39:41] == 'SK':
            shutil.move(file,  'C:/_____Passerelle_____/Slovakia/' + str(file[28:]))

        elif file[39:41] == 'RO':
            shutil.move(file, 'C:/_____Passerelle_____/Romania/' + str(file[28:]))

        elif file[39:41] == 'SP':
            shutil.move(file, 'C:/_____Passerelle_____/Spain/' + str(file[28:]))

        elif file[39:41] == 'PL':
            shutil.move(file, 'C:/_____Passerelle_____/Poland/' + str(file[28:]))

        elif file[39:41] == 'BE':
            shutil.move(file, 'C:/_____Passerelle_____/Belgium/' + str(file[28:]))

        elif file[39:41] == 'FR':
            shutil.move(file, 'C:/_____Passerelle_____/France/' + str(file[28:]))

time.sleep(15)
partition_files()
time.sleep(100)
partition_files()

print("Download complete !")
