import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import sys
import os
from os import listdir
from os.path import isfile, join
import pandas as pd
from pathlib import Path
import datetime
import pyautogui, sys
import psutil
from webdriver_manager.firefox import GeckoDriverManager
#enter your URL
urls = "https://sellercentral.amazon.com/reportcentral/INVENTORY_HEALTH/0"

#Enter your folder name where you want to save your File
file_path ='MIH'

#creating a funtion to get our profile directories
def get_profile_path(profile):
    FF_PROFILE_PATH = os.path.join(os.environ['APPDATA'],'Mozilla', 'Firefox', 'Profiles')

    try:
        profiles = os.listdir(FF_PROFILE_PATH)
    except WindowsError:
        print("Could not find profiles directory.")
        sys.exit(1)
    try:
        for folder in profiles:
            print(folder)
            if folder.endswith(profile):
                loc = folder
    except StopIteration:
        print("Firefox profile not found.")
        sys.exit(1)
    return os.path.join(FF_PROFILE_PATH, loc)


#Enter that folder path where you wanted to save your file
download_path_root = os.path.join(os.environ['USERPROFILE'],r'OneDrive\Reporting\U4')

#Firefox profiles
firefox_profile = 'Enter your profile name'


print("\n\nFetching Inventory MIH Report")

mime_types = r"text/csv"
profile = webdriver.FirefoxProfile(get_profile_path(firefox_profile))
profile.set_preference("browser.download.folderList", 2)
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", os.path.join(download_path_root, file_path))
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", mime_types)
profile.set_preference("plugin.disable_full_page_plugin_for_types", mime_types)

#Opening firefox browser
driver = webdriver.Firefox(firefox_profile=profile)
#maximizing browsers
driver.maximize_window()
wait = WebDriverWait(driver, 10)
#here the browser will enter our url 
driver.get(urls)
time.sleep(8)
#Clicking on "Downlaod"
driver.find_element(By.CSS_SELECTOR, '#reportpage_download_tab').click()
time.sleep(10)

#Requesting for latest csv file
driver.find_element(By.CSS_SELECTOR, 'kat-button[label = "Request .csv Download"]').click()
time.sleep(10)

#locating first row
table_row = driver.find_elements(By.CSS_SELECTOR, 'kat-table-body kat-table-row')
#locating first row's last column
table_cell = table_row[0].find_elements(By.CSS_SELECTOR, 'kat-table-cell')


#Creating a loop
while True:
    #checking if there is "In Progress" continue the loop
        if table_cell[4].text == 'In Progress':
            print('In Progress')
            time.sleep(60)
            continue
        #else click on the download button
        else:
            #try to find the download button
            try:
                download_btn = table_cell[4].find_element(By.CSS_SELECTOR, 'kat-button[label="Download"]')
                #if found click on the download button
                download_btn.click()
                #then break the loop and bot will be quit
                break
            #or else bot will wait for another 60 seconds until the bot generates the csv 
            except NoSuchElementException as exp :
                print("In Progress")
                time.sleep(60)
                continue

time.sleep(60)

driver.quit()
