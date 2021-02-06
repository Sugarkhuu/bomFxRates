"""
    
Requirements:
    Ms Excel - you can change to Libreoffice Calc or whatever you want if you know how to      
    Python   - 3 
    Selenium
    Chromedriver

Input: 
    Have this .py file (db_controller.py) and fx_template.xlsx in the current directory
    
Operation:
    Run this file, you will be prompted to enter "First Date" and "Last Date"
    for the period you want to collect the data. Lastly, it will ask
    wheter you want to skip weekends

Output:
    fx_created.xlsx will be created adding dates and rates to fx_template.xlsx 
    
Copyleft  Sugarkhuu Radnaa, Jan 2021    
    
"""

# import necessary packages
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options

import pandas as pd
from datetime import date, timedelta, datetime


# import fx_template
s_dir = "."
fx_list = pd.read_excel(s_dir+"//fx_template.xlsx",header=0)
fx_list['CC'] = fx_list['CC'].str.strip() # remove spaces from currency codes - is the case in this file which I copy/pasted from some source

# ask for user inputs: first and last dates; also whether to leave out weekends
firstDate = input("Insert First Date (Example: 2020-05-07): ")
lastDate  = input("Insert Last Date (Example: 2020-05-07): ")
skipWknd  = input("Skip weekend? If yes, type 'y', if not, type 'n'. (y/n): ")

# convert to datetime
firstDate = datetime.strptime(firstDate, '%Y-%m-%d')
lastDate  = datetime.strptime(lastDate, '%Y-%m-%d')
delta = timedelta(days=1)


# start chrome driver
chrome_options = Options()
chrome_options.add_argument("--headless") # comment it out if you want to see the Chrome on screen
driver = webdriver.Chrome(chrome_options=chrome_options)
bom_url_base = "https://www.mongolbank.mn/dblistofficialdailyrate.aspx"
driver.get(bom_url_base) # just testing if we can reach BoM FX daily page


# loop over fx data on BoM web
iDate = firstDate
while iDate <= lastDate:
    print ("Starting: ", iDate.strftime("%Y-%m-%d"))
    
    if iDate.isoweekday() in [6,7] and skipWknd == "y":
        print("A weekend! Jumping to next.")
        iDate += delta
        continue
    
    year = iDate.year
    month = iDate.month
    day = iDate.day
    
    # make url for a day
    bom_url_day = ''.join([bom_url_base,"?","vYear=",str(year),"&","vMonth=",str(month),"&","vDay=",str(day)])
    driver.get(bom_url_day)
        
    all_rates_path = "/html/body/form/main/div/div/div/div/div[2]/div/ul/li" # list of rates on that date
    all_rates = driver.find_elements_by_xpath(all_rates_path)
    
    for i in range(len(all_rates)-1):
        rate = all_rates[i].find_element_by_xpath("table/tbody/tr/td[3]/span").text # A fx rate
        rate_en = all_rates[i].find_element_by_xpath("table/tbody/tr/td[3]/span").get_attribute("id")[-3:] # FX code, 3 letters
        rate_mn = all_rates[i].find_element_by_xpath("table/tbody/tr/td[2]").text # MN long description of FX code
        fx_list.loc[fx_list['CC']==rate_en,iDate.strftime("%Y-%m-%d")] = float(rate.replace(',', '')) # putting rate to pd on the date
        # print(rate,rate_mn,rate_en, i)
    
    iDate += delta

# export to excel
fx_list.to_excel(s_dir+"//fx_created.xlsx", index=False)  
