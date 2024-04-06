import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

message = 'This is test message'
path = "test.xlsx"
wb_obj = openpyxl.load_workbook(path) 
sheet_obj = wb_obj.active

options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-extensions')
options.add_argument('--disable-logging')
options.add_argument('--disable-dev-shm-usage')
options.add_argument("--window-size=1366,768")
options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument('--ignore-certificate-errors')
options.add_argument("--password-store=basic")
options.add_argument("--enable-automation")
options.add_argument("--disable-browser-side-navigation")
options.add_argument("--disable-web-security")
options.add_argument("--disable-infobars")
options.add_argument("--disable-gpu")
options.add_argument("--disable-setuid-sandbox")
options.add_argument("--disable-software-rasterizer")

prefs = {"download.default_directory": "."}
options.add_experimental_option("prefs", prefs)
browser =  webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
action = ActionChains(browser)

BASE_URL = "https://web.whatsapp.com/"
n = 10

browser.get(BASE_URL)
print("Kindly go to output tab and login before the timeout of 120 seconds.")

# Wait for the learner to log in
try:
    wait = WebDriverWait(browser, 120)  # Adjust the timeout value as needed (e.g., 60 seconds)
    wait.until(EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']")))
    time.sleep(5)
    print("Logged in successfully. Proceeding to send message")
except Exception as e:
    print("Timeout: Learner did not log in within the specified time.")
    browser.quit()


for i in range(2,n):
    cell_obj = sheet_obj.cell(row=i, column=2)
    print(cell_obj.value)
    number = "+91" + str(int(cell_obj.value))

    browser.get("https://web.whatsapp.com/send?phone={}".format(str(number)))

    try:
        try:
            alert = browser.switch_to.alert
            print(alert.text)
            alert.accept()
        except Exception as e:
            print(e)

        inp_xpath = ('//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p')
        input_box = WebDriverWait(browser, 120).until(expected_conditions.presence_of_element_located((By.XPATH, inp_xpath)))

        input_box.send_keys(message)
        action.key_down(Keys.SHIFT).send_keys(Keys.ENTER).key_up(Keys.SHIFT).perform()
        input_box.send_keys(message)
        input_box.send_keys(Keys.ENTER)
        
        print("Message sent")
        time.sleep(2)

    except Exception as e:
        print("hey")

browser.quit()