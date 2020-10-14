from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import time
import json
import datetime as DT
import csv
import os
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


#browser.find_element_by_xpath
# from __future__ import print_function
# from desktop import logininfo
# print (data["username"])
# chromedriver = 'User/macbookowner/Desktop'
# use selenium/geckodriver to open chrome


def getCsv():
    today = DT.date.today()
    week_ago = today - DT.timedelta(days=7)


    with open('logininfo.json') as json_file:
        data = json.load(json_file)


    cc_url = 'https://www.e-zpassny.com/vector/account/home/accountLogin.do?locale=en_US&from=Home'


    browser = webdriver.Chrome()

    browser.get(cc_url)

    login= browser.find_element_by_name("login")
    password= browser.find_element_by_name("password")

    login.send_keys(data["username"])
    password.send_keys(data["userpass"])

    loginbtn= browser.find_element_by_name("btnLogin")
    loginbtn.click()
    # clicks the login button

    # --------Non used code----------##
    # button1= browser.find_element_by_link_text('Transactions')
    # button1.click()
    # button2= browser.find_element_by_link_text('&nbsp;&nbsp;•&nbsp;Transaction View')
    # button2.click()
    # print(browser.find_element_by_name('dateType').find_element_by_tag_name('option'))
    # ----------------------------------##

    browser.get('https://www.e-zpassny.com/vector/account/transactions/transactionSearch.do')
    browser.find_element_by_name('dateType').click()
    
    browser.find_element_by_xpath("//select[@name='dateType']/option[2]").click()

    strtdate= browser.find_element_by_name("startDateAll")
    strtdate.clear()
    week_ago_object = str(week_ago)[5:7] + '/' + str(week_ago)[8:10] + '/' + str(week_ago)[:4]
    strtdate.send_keys(week_ago_object)

    enddate1= browser.find_element_by_name("endDateAll")
    enddate1.clear()
    today_object = str(today)[5:7] + '/' + str(today)[8:10] + '/' + str(today)[:4]
    enddate1.send_keys(today_object)

    browser.find_element_by_name("btnSearch").click()

    browser.find_element_by_name("downloadType").click()
    browser.find_element_by_xpath("//select[@name='downloadType']/option[2]").click()

    browser.find_element_by_xpath("//img[@alt='Download']").click()
    time.sleep(10)
    browser.quit()


def updateGsheet():
    downloadSrc="/Users/macbookowner/Downloads"

    files = os.listdir(downloadSrc)
    paths = [os.path.join(downloadSrc, basename) for basename in files]
    ezPassExport = max(paths, key=os.path.getctime)
    ezPassRows=[]
    print(ezPassExport)
    with open (ezPassExport, newline='') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
        for row in spamreader:
            ezPassRows.append(row)

    if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        
    service = build('sheets', 'v4', credentials=creds)
    print(service)
    values = ezPassRows[1:]
    body = {
        'values': values
    }
    spreadsheetId='1-EeU7ViwhXndThqOFO0_WdzWhqlp53cRwyCnUq7KHoY'

    result = service.spreadsheets().values().append(
        spreadsheetId='1-EeU7ViwhXndThqOFO0_WdzWhqlp53cRwyCnUq7KHoY', range='Sheet2!A:Q',
        valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=body).execute()



    #------Non used code-------# 
    #  rowcount = service.spreadsheets().values().get(spreadsheetId, range).execute().getValues().size()
    # print(rowcount)
    # print('{0} cells append.'.format(result.get('updatedCells')))
    #---------------------------#




getCsv()
# getCsv runs the code to download the weekly ezpass to the computer

updateGsheet()
# updateGsheet extracts the download and uploads it to the driver's excel gsheet in the google API 