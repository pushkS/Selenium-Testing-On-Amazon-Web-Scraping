from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

from tkinter import *
import xlsxwriter 
import pandas as pd 
import xlrd
import requests
from bs4 import BeautifulSoup
import smtplib
import time

#selenium
PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)
driver.get("https://www.amazon.in")
print(driver.title)
driver.get("https://www.amazon.in/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.in%2F%3Fref_%3Dnav_signin&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=inflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&")

#login
driver.find_element_by_id('ap_email').send_keys('8308451787')
driver.find_element_by_id('continue').click()
driver.find_element_by_id('ap_password').send_keys('pushkR@777')
driver.find_element_by_id('signInSubmit').click()

prod = driver.find_element_by_name("field-keywords")
prod.send_keys("MSI Gaming GeForce RTX 2060 6GB GDRR6")
prod.send_keys(Keys.RETURN)
driver.get("https://www.amazon.in/dp/B07NJ4N5NV/ref=cm_sw_em_r_mt_dp_U_EKvVDbCYVFFR6")
try:
    ppd = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "ppd"))
    )
    time.sleep(5)
    print(ppd.text)
except:
    driver.quit()

driver.quit()

#main
window = Tk()
window.geometry('480x360')
window.title("Web scrapping")

lbl0 = Label(window, text="Amazon Price Monitoring",font='Helvetica 18 bold')
lbl0.grid(column=1, row=0)

lbl = Label(window, text="URL")
lbl.grid(column=0, row=1)
txt = Entry(window,width=30)
txt.grid(column=1, row=1)

lbl2 = Label(window, text="Email id")
lbl2.grid(column=0, row=5)
txt2 = Entry(window,width=30)
txt2.grid(column=1, row=5)

lbl3 = Label(window, text="Price to be compared")
lbl3.grid(column=0, row=9)
txt3 = Entry(window,width=30)
txt3.grid(column=1, row=9)
num_array = []
def clicked():
    n = txt.get()
    u = txt2.get()
    p = txt3.get()

    num_array.append([n,u,int(p)])
    print (num_array)

btn = Button(window, text="Add",command=clicked)
btn.grid(column=1, row=13)

def submit():
    df = pd.DataFrame(num_array,columns =['URL','Email','Price'])

    df.to_excel('result.xlsx', index = False)

btn1 = Button(window, text="Save to Excel",command=submit)
btn1.grid(column=1, row=17)

pass1="reaxykwydcvghqmj"
def exe():
    loc = ("result.xlsx")

    # To open Workbook 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0) 

    sheet.cell_value(0, 0) 
    row_count = sheet.nrows
    cnt =1 
    while cnt <= row_count:
        xs = sheet.row_values(cnt)

        URL1=xs[0]
        ma1=xs[1]
        pri1=xs[2]
        print (URL1)
        print("\nHELLO WELCOME TO ONLINE PRICE TRACKER:-")

        headers={"User-Agent":'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36'}

        def check_price():
            page = requests.get(URL1, headers=headers)
            soup=BeautifulSoup(page.content, 'html.parser')
            title = soup.find(id="productTitle").get_text()
            print(title.strip())
            price =soup.find(id="priceblock_ourprice").get_text()
            converted_price = float(price[2:4])
            if(converted_price<pri1):
                send_mail(URL1,ma1,pri1)
            print(converted_price)

            
        check_price()
        cnt += 1
            
def send_mail(URL1, ma1, pri1):
    # pri=int(txt3.get())
    ma = "pushkarajsonawane1998@gmail.com" 
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(ma, pass1)
    subject = 'Price fell down!!!'
    body = 'Check the amazon link: ' + URL1
    msg = f"Subject: {subject}\n\n{body}"
    server.sendmail(ma,ma1,msg)
    print('\nHey Email has been sent!!')
    server.quit()

btn2 = Button(window, text="execute script",command=exe)
btn2.grid(column=1, row=21)

window.mainloop()
# https://www.amazon.in/dp/B07NJ4N5NV/ref=cm_sw_em_r_mt_dp_U_EKvVDbCYVFFR6

#https://www.amazon.in/dp/B078BNQ318/ref=cm_sw_em_r_mt_dp_7yXZFbQTT36BG
#https://www.amazon.in/s?k=OnePlus+8&i=electronics&ref=nb_sb_noss_2     search bar
