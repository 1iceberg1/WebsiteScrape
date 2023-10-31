import sys
import os
from openpyxl.utils import exceptions
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time
import sched
from Analyzer import Analyzer
from Record import Record
import re
import datetime

UserName = 'nextaa'
PassWord = 'Qwer1234'

def check_file_writable(file_path):
    if not os.path.exists(file_path):
        return False
    try:
        # Check if the file is open
        with open(file_path, 'r') as file:
            pass  # File is not open or read-only
    except PermissionError:
        return False  # File is open or read-only
    except FileNotFoundError:
        return False  # File does not exist
    return True  # File is not open or read-only

def get_date_value(page):
    soup = BeautifulSoup(page, 'html.parser')

    date_value = str(soup.find('span', {'id' : 'timecontainer'}).text)
    date_value = "".join(date_value.split())

    date_value = date_value[:-5]
    if len(date_value) == 18:
        if date_value[1].isdigit():
            if date_value[9] == '0':
                date_value = date_value[:9] + "12" + date_value[10:]
            else:
                date_value = date_value[:9] + "0" + date_value[9:]
        else:
            date_value = "0" + date_value
    elif len(date_value) == 17:
        if date_value[8] =='0':
            date_value = "0" + date_value[:8] + "12" + date_value[9:]
        else:
            date_value = "0" + date_value[:8] + "0" + date_value[8:]

    input_format = "%d%b%Y%I:%M:%S%p"
    output_format = "%m/%d/%Y, %I:%M:%S%p"

    date_obj = datetime.datetime.strptime(date_value, input_format)
    formatted_date_str = date_obj.strftime(output_format)
    
    return formatted_date_str

def calculate_passed_minutes(time_str):
    # Split the time string into hours, minutes, seconds, and AM/PM components
    print("TIME: " + time_str)
    time_parts = time_str[:-2].split(":")
    hours = int(time_parts[0])
    minutes = int(time_parts[1])
    seconds = int(time_parts[2])
    am_pm = time_str[-2:]

    # Adjust hours if it's PM
    if am_pm.lower() == "pm":
        hours += 12

    # Calculate the total number of minutes
    total_minutes = (hours * 60) + minutes + (seconds // 60)
    
    return total_minutes

def save_workbook(filename, workbook):
    success_flag = True
    try:
        workbook.save(filename)
        print("Successfully Recorded! Please check the file.")
    except PermissionError:
        print("Failed to save the workbook. The file is currently open.")
        success_flag = False
    except exceptions.ReadOnlyWorkbookException:
        print("Failed to save the workbook. The file is opened as read-only.")
        success_flag = False
    except Exception as e:
        print("An error occurred while saving the workbook:", str(e))
        success_flag = False

    if success_flag == False:
        print("Please check your file is currently open or opend as read-only")

def ScrapeData():
    
    start_time = time.time()
    
    url = "https://www.m8clicks.com"
    early_url = "https://m8clicks.com/_View/RMOdds2.aspx?ot=e&ov=1&mt=0&wd=&isWC=False&ia=2&tf=-1"
    # today_url = "https://m8clicks.com/_View/RMOdds2.aspx?ot=t&ov=1&isWC=False&ia=0&isSiteFav=False"
    today_url = "https://m8clicks.com/_View/RMOdds2.aspx?ot=t&ov=1&mt=0&wd=&isWC=False&ia=0&tf=-1"

    if not os.path.exists('record.xlsx'):
        workbook = Workbook()
        save_workbook('record.xlsx', workbook)

    while check_file_writable("record.xlsx") != True:
        print("Please close opened Excel file: record.xlsx")
        time.sleep(5)

    # # specify  the user profile directory
    # user_profile_directory = 'C:/selenium'

    # # create ChromeOptions object and set the user profile directory
    # chrome_options = Options()
    # chrome_options.add_experimental_option("detach", True)
    # chrome_options.add_argument("user-data-dir=" + user_profile_directory)
    # chrome_options.add_argument('--profile-directory=selenium')
    # chrome_options.add_argument('--incognito')
    # chrome_options.add_argument("--enable-javascript")
    # chrome_options.add_argument("--enable-file-cookies")
    # chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
    # driver = webdriver.Chrome(options = chrome_options)

    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_argument('--headless')
    chromeOptions.add_argument('--ignore-certificate-errors')

    # instantiate a webdriver
    driver = webdriver.Chrome(options = chromeOptions)

    # get date
    driver.get("https://www.m8clicks.com/ClockTime.aspx")
    time.sleep(3)
    driver.execute_script("return document.readyState")
    time.sleep(2)

    date_html = driver.page_source
    # print(date_html)
    current_date_time = get_date_value(date_html)
    passed_minutes = calculate_passed_minutes(current_date_time[-10:])
    current_date = current_date_time[:10]

    print(current_date)
    print(passed_minutes)

    # First Login

    # if not os.path.exists("cookies.pkl"):
    driver.get(url)
    time.sleep(3)
    username_filed = driver.find_element('id', 'txtUserName')
    password_filed = driver.find_element('id', 'txtPassword')

    username_filed.send_keys(UserName)
    password_filed.send_keys(PassWord)
    
    password_filed.send_keys(Keys.RETURN)

    time.sleep(2)

    # After logging in
    cookies = driver.get_cookies()

    # Store the cookies
    cookie_dict = {}
    for cookie in cookies:
        cookie_dict[cookie['name']] = cookie['value']

    time.sleep(5)

    while url == driver.current_url:
        time.sleep(1)

    print(driver.current_url)

    current_url = driver.current_url

    current_url = re.sub(r"lang=[A-Z]{2}-[A-Z]{2}", "lang=EN-US", current_url)
    
    driver.get(current_url)
    time.sleep(2)

    # Today Recording

    try:
        driver.get(today_url)

        # Add the stored cookies to the WebDriver
        for name, value in cookie_dict.items():
            driver.add_cookie({'name': name, 'value': value})

        driver.refresh()

    except Exception as e:
        print("Error: " + str(e))
    
    
    # this is just to ensure that the page is loaded 
    time.sleep(5) 
    
    driver.execute_script("return document.readyState")

    
    time.sleep(2)

    html = driver.page_source 
    # print(html)

    analyzer = Analyzer(html, True, current_date)
    leagues = analyzer.get_data()

    record = Record(current_date, passed_minutes)
    record.set_data(leagues)
    record.create_file_sheet("record.xlsx")

    # Early Recording

    try:
        driver.get(early_url)

        # Add the stored cookies to the WebDriver
        for name, value in cookie_dict.items():
            driver.add_cookie({'name': name, 'value': value})

        driver.refresh()

    except Exception as e:
        print("Error: " + str(e))

    # this is just to ensure that the page is loaded 
    time.sleep(10) 

    html = driver.page_source 
    # print(html)
    
    driver.execute_script("return document.readyState")

    time.sleep(2)

    analyzer = Analyzer(html, False, current_date)
    leagues = analyzer.get_data()

    record = Record(current_date, passed_minutes)
    record.set_data(leagues)
    record.create_file_sheet("record.xlsx")

    print("URL")
    print(driver.current_url)

    driver.close() # closing the webdriver 

    end_time = time.time()

    return (end_time - start_time, passed_minutes)


# ScrapeData()

def call_function(scheduler):

    runtime, record_time = ScrapeData()
    print("RECORD TIME: " + str(record_time))
    print("Runtime: " + str(runtime))
    # Calculate
    current_seconds = record_time * 60 + int(runtime) + 1
    next_time = current_seconds // 1800 + 1
    after = 1800 * next_time - current_seconds
    if after < 0: after = 0
    print("Record after " + str(after) + " seconds again")
    scheduler.enter(after, 1, call_function, (scheduler,))

my_scheduler = sched.scheduler(time.time, time.sleep)
my_scheduler.enter(1, 1, call_function, (my_scheduler,))
my_scheduler.run()