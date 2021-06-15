import openpyxl
import datetime
import time
import subprocess
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_argument("--headless")

# Creating today's date in format dd.mm.yyyy

time_now = datetime.datetime.now()
day = time_now.strftime('%d')
mon = time_now.strftime('%m')
year = time_now.strftime('%Y')
todayDate = (day+"."+mon+"."+year)

# Opening excel file and sheet

wb = openpyxl.load_workbook(r'C:\Users\XXXXX\XXXXX\Report.xlsx')
sheet = wb['2021']

# Taking first and last rows numbers from user

n = int(input("Type the number of the first row: "))
m = int(input("Type the number of the last row: "))

# Loop for rows from n to m

while n <= m:
    n_eb = ("H"+str(n))  # early bird or issue cell
    n_art = ("G"+str(n))  # journal articles number cell
    n_date = ("M"+str(n))  # cell for date printing
    n_code = ("D"+str(n))  # journal code cell
    n_upload = ("L"+str(n))  # upload date cell
    art = str(sheet[n_art].value)
    code = str(sheet[n_code].value).strip()
    if_eb = str(sheet[n_eb].value)
    date = str(sheet[n_date].value)
    print(code)
    
    if date == "None" and code != "None":

        if if_eb == "EB":
            url = ("https://XXXXX.com/issue/"+code+"/0/0")  # eb url creating from above data
            browser = webdriver.Chrome(ChromeDriverManager().install(), options=options)
            browser.get(url)
            time.sleep(3)
        
            elem = browser.find_element_by_xpath('//*[@id="article-list-pc"]/div[1]/span').text  # finding number of articles on url webpage
            artNum = int(elem.split()[0])

            i = 1
            thisEB = 0

            while i <= artNum:
                EB_Date_Path = ('//*[@id="article-list-pc"]/div[2]/div['+str(i)+']/div[1]/span[1]')
                pubOnl = browser.find_element_by_xpath(EB_Date_Path).text  # finding published online date
                pubOnldate = datetime.datetime.strptime(pubOnl[-11:], '%d %b %Y')
                eb_d = pubOnldate.strftime('%Y-%m-%d')
                
                uplDate = str(sheet[n_upload].value)
                
                if eb_d == uplDate[0:10]:
                    thisEB += 1  # counting articles online from this upload

                i += 1

            if int(art) == thisEB:
                sheet[n_date] = todayDate # printing today's date to cell
            else:
                sheet[n_date] = "error"

        else:    
            n_vol = ("E"+str(n))  # journal volume cell 
            n_iss = ("F"+str(n))  # journal issue cell
            n_cover = ("P"+str(n))
            vol = str(sheet[n_vol].value)
            iss = str(sheet[n_iss].value)
            url = ("https://XXXXX.com/issue/"+code+"/"+vol+"/"+iss)  # issue url creating from above data 
            browser = webdriver.Chrome(ChromeDriverManager().install(), options=options)
            browser.get(url)
            time.sleep(3)
            elem = browser.find_element_by_xpath('//*[@id="article-list-pc"]/div[1]/span').text  # finding number of articles on url webpage
                   
            numArt = (art+" Articles")
            if numArt == elem:
                sheet[n_date] = todayDate # printing today's date to cell
            else:
                sheet[n_date] = "error"
                
        wb.save(r'C:\Users\XXXXX\XXXXX\Report.xlsx')
        browser.quit()

    n += 1
    
# final saving and opening the file
    
wb.save(r'C:\Users\XXXXX\XXXXX\Report.xlsx')
subprocess.Popen(['start', r'C:\Users\XXXXX\XXXXX\Report.xlsx'], shell=True)
