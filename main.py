import time
import pyautogui
import xlrd
from selenium import webdriver

with open("config.txt") as f:
    path = f.readline()
    loc = ((next(f)))

driver_path = "chromedriver.exe"  #download a chrome driver from https://chromedriver.chromium.org/downloads and set its path here
options = webdriver.ChromeOptions() 
options.add_argument(str(path)) #path to your locally stored chrome profile
driver = webdriver.Chrome(driver_path, options = options)
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
n = 10 #number of messages you want to send

def new_line():
        pyautogui.keyDown('shift')
        pyautogui.press('enter')
        pyautogui.keyUp('shift')

for i in range(1,n):
        msg =  "hi"
        number = "+91" + str(int(sheet.cell_value(i,1)))
        print(i)
        driver.get("https://web.whatsapp.com/send?phone={}".format(str(number)))

        time.sleep(15)
        pyautogui.moveTo(600, 980) #pixels values are subject to change based on your screen
        pyautogui.click()

        pyautogui.write(msg)
        new_line() 

        pyautogui.moveTo(920, 980)
        pyautogui.click()

        time.sleep(3)