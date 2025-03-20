from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys

from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.chrome.service import Service

import time
import os
import sys
import random
import string
import re

import pandas as pd

from string import *
from random import randint
from selenium.common.exceptions import *
from os import path

from datetime import date, timedelta, datetime
import glob
import shutil

#! driver = webdriver.Chrome("C:\webdrivers\chromedriver.exe")

import chromedriver_autoinstaller
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service = service)

#chrome_options = Options()
#chrome_options.set_headless()
#driver =webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chrome_options)

#driver.get("https://www.e2open.com/about")
driver.get("https://orange.e2open.com/ORANGE_na/e2na/console/web/main.action")
driver.maximize_window()

wait = WebDriverWait(driver, 8)
#log_in = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@href='#']")))

actions = ActionChains(driver)
#actions.move_to_element(log_in).perform()
time.sleep(2)

#e2open = driver.find_element_by_xpath("//a[contains(@href, 'network.e2open')]")
#actions.move_to_element(e2open).click().perform()
'''
original_window = driver.current_window_handle
#assert len(driver.window_handles) == 1
wait.until(EC.number_of_windows_to_be(2))

for window_handle in driver.window_handles:
    if window_handle != original_window:
        driver.switch_to.window(window_handle)
        break

cookie = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//button[@id = 'accept-button']")))
cookie.click()
'''

##################################
username = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='username']")))
username.send_keys('fabrice.francois@orange.com') # <-----------------------

password = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='password']")))
password.send_keys('Thursday@123*-') #<------------------
##################################

submit = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name = 'submit']")))
submit.click()
#driver.find_element_by_xpath("//input[@name = 'submit']").click()

#wait.until(EC.presence_of_element_located((By.XPATH, "//img[@alt = 'Network Console']"))).click()

def change_date(day):
    global iframe
    time.sleep(3)
    iframe = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//iframe[@src='https://orange.e2open.com/ORANGE_na/e2na/console/servlet?request_type=console/transaction/search']")))
    #iframe = driver.find_element_by_xpath("//iframe[@src='https://orange.e2open.com/ORANGE_na/e2na/console/servlet?request_type=console/transaction/search']")
    driver.switch_to.frame(iframe)

    #wait.until(EC.presence_of_element_located((By.XPATH, "//input[@class='buttonRegular']"))).click()
    date_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name = 'e2startdate']")))
    date_input.click()
    date_input.clear()
    date_input.send_keys(day.strftime('%d-%m-%Y'))

    state = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//select[@name = 'e2state']")))
    #state = driver.find_element_by_xpath("//select[@name = 'e2state']")
    state.click()
    time.sleep(3)
    state.send_keys(Keys.END)

    warning = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//option[@value = 'Warnings']")))
    warning.click()
    #driver.find_element_by_xpath("//option[@value = 'Warnings']").click()

    """ActionChains(driver).key_down(Keys.CONTROL).\
        click("//option[@value = 'Warnings']").\
        click("//option[@value = 'Errors']").\
        key_up(Keys.CONTROL).perform()
    """
    try:
        ActionChains(driver).send_keys(Keys.SHIFT, Keys.ARROW_DOWN).perform()
    except StaleElementReferenceException as e:
        pass

jour = datetime.now().strftime("%a")
if jour == 'Mon':
    day_3back = date.today() - timedelta(days=3)
    print(day_3back)
    change_date(day_3back)
else:
    yesterday = date.today() - timedelta(days=1)
    #print(yesterday.strftime('%d-%m-%Y'))
    change_date(yesterday)

WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@class = 'buttonRegular' and @value = 'Search']"))).click()
#driver.find_element_by_xpath("//input[@class = 'buttonRegular' and @value = 'Search']").click()

WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@class = 'buttonRegular' and @value = 'Export']"))).click()
#driver.find_element_by_xpath("//input[@class = 'buttonRegular' and @value = 'Export']").click()

time.sleep(2)
transactions = pd.read_csv("C:/Users/"+ str(os.getlogin()) +"/Downloads/transactions.csv")

#transactions1 = transactions1[~transactions1['Message'].str.contains('Shipment')]
transactions = transactions[transactions['Message'].str.contains('DiscreteOrder - 1.0')]

trans_id = transactions['Transaction ID'].tolist()

#print(trans_id)
driver.back()
driver.switch_to.frame(iframe)

directory = 'C:/Users/'+ str(os.getlogin()) +'/Downloads/'
#latest_file = max(glob.glob(directory + 'transactions (*).csv'), key = os.path.getctime)

#for filename in os.listdir(directory):
#    if  os.path.getsize(latest_file) >= 200:
#        print(filename)

#newpath = 'C:/Users//Desktop/Errors_'+ str(datetime.now().strftime("%a"))
newpath = 'C:/Errors_'+ str(datetime.now().strftime("%a"))
if not os.path.exists(newpath):
    os.makedirs(newpath)

elif os.path.exists(newpath):
    for filename in os.listdir(newpath):
        if os.path.isfile(os.path.join(newpath, filename)):
            os.remove(os.path.join(newpath, filename))
#i = 1
for i in range(len(trans_id)):
    print(trans_id[i], i)

    tr_id = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name = 'e2transactionId']")))
    tr_id.send_keys(trans_id[i])
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@value = 'Search']"))).click()
    #driver.find_element_by_xpath("//input[@value = 'Search']").click()
    #wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value = 'Export']"))).click()
    wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'javascript:selectTransaction')]"))).click()
    time.sleep(3)

    table_events = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[@class='tableBorder']/tbody/tr")))
    print(len(table_events))
    #last_event = driver.find_element_by_xpath("//tr[" + str(len(table_events)) + "]//a")
    last_event = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//tr[" + str(len(table_events)) + "]//a")))

    print(last_event.text)
    last_event.click()
    
    time.sleep(1)

    ##!!wait.until(EC.presence_of_element_located((By.XPATH, "//tbody/tr[15]/td/a[contains(@href, 'javascript:selectTransaction')]"))).click()
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//td[3]/a[2][contains(@href, 'javascript:processTransaction')]"))).click()
    
    except TimeoutException as e:
        #driver.find_element_by_xpath()
        print("Status ERROR")
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//a[1][contains(@href, 'javascript:processTransaction')]"))).click()
        except NoSuchElementException as e1:
            print("-----> This is a one-time error with ID: ", trans_id[i])

    except StaleElementReferenceException as e2:
        time.sleep(2)
        pass
        
    time.sleep(3)

    driver.back()
    driver.back()
    driver.back()
    time.sleep(2)

    '''files1 = glob.glob('C:/Users//Downloads/*.file') # <---
    #print(files)

    for file_1 in files1:
        new_name = f'C:/Users//Downloads/'+ str(trans_id[i]) + '.file'
        new_name1 = 'C:/Users//Downloads/'+ str(trans_id[i]) + '.file'
        try:
            os.rename(file_1, new_name)
        except Exception as e:
            pass
        file_1 = open(new_name, 'r')
        f = file_1.read()
        uniqueWords=set(f.split())
        if os.path.exists('C:/Users//Downloads/'+ str(trans_id) +'.txt'):
            print("The ID exists aleady")
            break
        elif 'Supplier' in uniqueWords:
            try:
                shutil.move(new_name1, newpath + '/' + str(trans_id[i]) + '.file')
                #shutil.move(new_name1, newpath)
            except Exception as e:
                pass
    #driver.switch_to.frame(iframe)
    #driver.find_element_by_xpath("//input[@name = 'e2transactionId']").clear()
    i += 1 # <--- trial '''

    ''' 18.03 edit
    # Loop through each file and rename them
    files = glob.glob('C:/Users//Downloads/discrete*.err')
    for file in files:
        # Split the file name at the '_' character
        #parts = file.split('_')
        #print(file)

        # Create a new file name using the parts and the desired prefix
        new_name = f'C:/Users//Downloads/'+ str(trans_id[i]) + '.txt'
        new_name1 = 'C:/Users//Downloads/'+ str(trans_id[i]) + '.txt'
        # Rename the file using os.rename()
        os.rename(file, new_name)

        file = open(new_name, 'r')
#
        f = file.read()
        def move():
            try:
                shutil.move(new_name1, newpath + '/' + str(trans_id[i]) + '.txt')
                #shutil.move(new_name1, newpath)
            except PermissionError as e:
                pass

        uniqueWords=set(f.split())
        
        #if 'C:/Users//Downloads/'+ str(trans_id) +'.file' in 'C:/Users//Downloads/':
        if os.path.exists('C:/Users//Downloads/'+ str(trans_id) +'.file'):
            break
        elif 'Promise' in uniqueWords:
            move()
        else:
            pass

    18.03 '''
    files = glob.glob('C:/Users/'+ str(os.getlogin()) +'/Downloads/*.*')

    for file in files:
        def move():
            try:
                shutil.move(new_name, newpath + '/' + str(trans_id[i]) + '.txt')
                #shutil.move(new_name1, newpath)
            except Exception as e:
                pass
        
        if file.endswith('discreteOrder.err'):
            new_name = f'C:/Users/'+ str(os.getlogin()) +'/Downloads/'+ str(trans_id[i]) + '.txt'
            # Rename the file using os.rename()
            os.rename(file, new_name)
            file = open(new_name, 'r')
            f = file.read()

            uniqueWords=set(f.split())
        
            if 'Promise' in uniqueWords:
                move()
            else:
                pass
        #elif file.__contains__('.file'):
        elif file.endswith('.file'):
            print("TEST", file)
            new_name = f'C:/Users/'+ str(os.getlogin()) +'/Downloads/'+ str(trans_id[i]) + '.txt'
            print('extension .file has id: ', trans_id[i])
            # Rename the file using os.rename()
            #time.sleep(3)
            try:
                os.rename(file, new_name)
            except PermissionError as e:
                time.sleep(3)
                os.rename(file, new_name)
            except FileNotFoundError as e:
                os.rename(file, new_name)
            file = open(new_name, 'r')
            f = file.read()

            uniqueWords=set(f.split())
            #print(uniqueWords)
        
            if 'Header' in uniqueWords:
                move()
            else:
                pass
        else:
            pass

    driver.switch_to.frame(iframe)
    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name = 'e2transactionId']"))).clear()

    #driver.find_element_by_xpath("//input[@name = 'e2transactionId']").clear()
    i += 1

print("Download is complete!")