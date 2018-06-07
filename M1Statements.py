from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
import time
import PyPDF2
import os
import glob
import re
import csv

#Set the download path and initialize the chrome driver
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : "C:\Python34\M1InvestStatements"}
chromeOptions.add_experimental_option("prefs",prefs)
chromedriver = r"C:\Users\David\Downloads\chromedriver.exe"
driver = webdriver.Chrome(executable_path=chromedriver, chrome_options=chromeOptions)

#Go to M1 website
url = "https://dashboard.m1finance.com/d/activity/trade-confirmations"
driver.get(url)
time.sleep(10)

#Log in 
username = ""
password = ""

usernameField = driver.find_element_by_xpath("""//*[@id="root"]/div/div/div[2]/div[3]/div/form/div[1]/div/div[1]/div/input""")
usernameField.send_keys(username)
pwField = driver.find_element_by_xpath("""//*[@id="root"]/div/div/div[2]/div[3]/div/form/div[1]/div/div[2]/div/input""")
pwField.send_keys(password)
driver.find_element_by_xpath("""//*[@id="root"]/div/div/div[2]/div[3]/div/form/div[3]/div/div/button/div/span[2]""").click()
time.sleep(10)

#Delete old PDFs
for i in range(1,20):
    try:
        os.remove('C:\Python34\M1InvestStatements\statement' + str(i) + '.pdf')
    except FileNotFoundError:
        print("No such file")

#Download statement PDFs
PDFcount = 0
for i in range(1,20):
    try:
        driver.find_element_by_xpath("""//*[@id="root"]/div/div/div[1]/div[2]/div/div/div[2]/div/table/tbody/tr[""" + str(i) + """]/td[2]/a/span""").click()
        time.sleep(15)
        #Rename the downloaded PDF
        list_of_files = glob.glob('C:\Python34\M1InvestStatements/*') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getctime)
        os.rename(latest_file, 'C:\Python34\M1InvestStatements\statement' + str(i) + '.pdf')
        PDFcount+=1
    except WebDriverException:
        print("No More")

##########PDF Work############ 
    
output_file = open('M1Stocks.csv', 'w')
output_file.write('QTY,SYM,PRICE PAID,TOTAL PAID,,SYMS,CURRENT PRICE')
uniqueSyms = set([])
uniquePrices = []
uniqueSyms2 = []

#Convert PDFs to txt
for i in range(1,PDFcount+1):
    with open('statement' + str(i) + '.pdf','rb') as pdf_file, open('PDFtext.txt', 'w') as text_file:
        read_pdf = PyPDF2.PdfFileReader(pdf_file)
        number_of_pages = read_pdf.getNumPages()
        for page_number in range(number_of_pages):   # use xrange in Py2
            page = read_pdf.getPage(page_number)
            page_content = page.extractText()
            text_file.write(page_content)


    #date,date,QTY,SYM,PRICE

    input_file = open('PDFtext.txt', 'r')
  #Using regex and other tools to parse out the relevant data (QTY, SYM, PRICE)  
    consDates = 0
    attribs = 3
    newStock = ""
    for line in input_file:
        if consDates == 2:
            if attribs == 2:
                uniqueSyms.add(line.replace('\n', '').replace('\r', ''))
                if line.replace('\n', '').replace('\r', '') not in uniqueSyms2:
                    uniqueSyms2.append(line.replace('\n', '').replace('\r', ''))
            if attribs > 0:
                if attribs == 3:
                    newStock = newStock + line
                    attribs-=1
                else:
                    newStock = newStock + "," + line
                    attribs-=1
            else:
                 output_file.write('\n' + newStock.replace('\n', '').replace('\r', ''))# + ',,')
                 print(newStock.replace('\n', '').replace('\r', ''))
                 newStock = ""
                 attribs = 3
                 consDates = 0
            
        if re.match(r'(\d+/\d+/\d+)', line) is not None:
            consDates+=1
            
output_file.close()
print(uniqueSyms2)  
#print (uniquePrices[0])


#Go get current stock prices
url = "http://bigcharts.marketwatch.com/quotes/multi.asp?view=Q"
driver.get(url)
time.sleep(10)

quotes = driver.find_element_by_xpath("""//*[@id="msymb"]""")
for sym in uniqueSyms2:
    quotes.send_keys(sym + " ")
driver.find_element_by_xpath("""//*[@id="quickquotes"]/form/div/div[2]/div[1]/input[2]""").click()
time.sleep(10)

for i in range(1,len(uniqueSyms2)+1):
    uniquePrices.append(driver.find_element_by_xpath("""//*[@id="quickquotes"]/table/tbody/tr[""" + str(i) + """]/td[2]""").text)

#Add a column for unique stock symbols and prices
with open('M1Stocks.csv', 'r') as f:
    reader = csv.reader(f)
    data = [next(reader)]
    for row in reader:
        if str(row[1]) in uniqueSyms:
            data.append(row + [float(row[0])*float(row[2])] + [" "] + [row[1]] + [uniquePrices[0]])
            uniqueSyms.remove(str(row[1]))
            del uniquePrices[0]
        else:
            data.append(row + [float(row[0])*float(row[2])] + [" "])
            
with open('M1Stocks.csv', 'w', newline='') as nf:
    writer = csv.writer(nf)
    writer.writerows(data)
    
print(uniqueSyms)
