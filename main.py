import sys
import os
from openpyxl.utils import exceptions
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time
import sched
from Analyzer import Analyzer
from Record import Record
import re
import datetime
from selenium.common.exceptions import NoSuchElementException

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

def change_date_format(date_str, day_diff, input_format, output_format):
    date_obj = datetime.datetime.strptime(date_str, input_format)
    date_obj = date_obj + datetime.timedelta(days = day_diff)
    output_date = date_obj.strftime(output_format)
    return output_date

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

    if hours == 12: hours = 0

    # Adjust hours if it's PM
    if am_pm.lower() == "pm":
        hours += 12

    # Calculate the total number of minutes
    total_minutes = (hours * 60) + minutes + (seconds // 60)
    
    return total_minutes

def ScrapeData():
    
    start_time = time.time()
    
    url = "https://www.m8clicks.com"
    early_url1 = "https://m8clicks.com/_View/RMOdds2.aspx?ot=e&ov=1&mt=0&wd="
    # early_url1 = "https://m8clicks.com/_View/RMOdds2.aspx?update=false&r=316466324&wd="
    # "ot=e&ov=1&isWC=False&ia=0&isSiteFav=False"
    early_url2 = "&isWC=False&ia=0&tf=-1"
    today_url = "https://m8clicks.com/_View/RMOdds2.aspx?ot=t&ov=1&mt=0&wd=&isWC=False&ia=0&tf=-1&isSiteFav=False"
    # "https://m8clicks.com/_View/RMOdds2.aspx?update=false&r=316466324&wd=2023-10-31&ot=e&isWC=False&ia=0&isSiteFav=False"

    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_argument('--headless')
    chromeOptions.add_argument('--allow-running-insecure-content')
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
    
    try:

        username_filed = driver.find_element('id', 'txtUserName')
        password_filed = driver.find_element('id', 'txtPassword')

        username_filed.send_keys(UserName)
        password_filed.send_keys(PassWord)
        
        password_filed.send_keys(Keys.RETURN)
    
    except NoSuchElementException:
        
        driver.close() # closing the webdriver 
        print("The website is not available now!")
        print("Please try again later...")
        
        end_time = time.time()

        return (end_time - start_time, passed_minutes)

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

    analyzer = Analyzer(html, True, current_date, passed_minutes)
    leagues = analyzer.get_data()

    if len(leagues) != 0:
        record = Record(current_date, passed_minutes)
        record.set_data(leagues)
        record.create_file_sheet("record.xlsx")

    # Early Recording

    k = 0
    
    for i in range(6):
        early_date_format = change_date_format(current_date, i, "%m/%d/%Y", "%Y-%m-%d")
        early_url = early_url1 + early_date_format + early_url2

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

        analyzer = Analyzer(html, False, current_date, passed_minutes)
        leagues = analyzer.get_data()

        if len(leagues) != 0:
            record = Record(current_date, passed_minutes)
            record.set_data(leagues)
            record.create_file_sheet("record.xlsx")
            k = k + 1
        if k == 5: break


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