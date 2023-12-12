import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from openpyxl import Workbook, load_workbook 
import pandas

wb=load_workbook('salary.xlsx')
sheet=wb['result']

service = Service()
option = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=option)
total=0
j=1
name=[]
fails = pandas.read_excel("salary.xlsx", sheet_name="Sheet2")

info_list=fails.values.tolist()

url = "https://emn178.github.io/online-tools/crc32.html"
driver.get(url)
time.sleep(1)
find=driver.find_element(By.ID,"input")
# program read information from people.csv file and put all data in name list.
with open("people.csv", "r") as file:
    next(file)
    for line in file:
        row=line.rstrip().split(",") 
        name=row[2]+" "+row[3]
        find.clear()
        find.send_keys(name)
        find2=driver.find_element(By.ID,"output")
        find3=find2.get_attribute("value")
        for x in range(len(info_list)):
            if(info_list[x][0]==find3):
                total=total+info_list[x][1]
                
        sheet['A'+str(j)]=name
        sheet['B'+str(j)]=total
        j=j+1
        print(name)
        print(total)
        total=0
wb.save("salary.xlsx")
 

