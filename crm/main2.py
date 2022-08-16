from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
import openpyxl
import unicodedata
from defs2 import insertCustomerInfo
from defs2 import getElementsInfo
from defs2 import getCustomerInfo
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

capabilities = DesiredCapabilities.CHROME.copy()
capabilities['acceptInsecureCerts'] = True
##########################################################################################################
######
#設定#
######
url = "http://entrydev.examobile.jp"
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome("chromedriver.exe", desired_capabilities=capabilities)
driver.get(url)
#identify挿入画像
image_file = ("img\\chrome_ver.png")
#エクセル読み取りファイル指定
book = openpyxl.load_workbook("customer_info.xlsx", data_only=True)
sheet = book["お客様情報2"]
num_of_customers = int(sheet["C1"].value)

start_row = 6
start_column = 3

for i in range(num_of_customers):
    
    #excelから顧客情報を読み取る
    customer_info = getCustomerInfo(start_row, start_column, sheet)
    #未入力項目(None)を空白に変換
    for key in customer_info:
        if customer_info[key] is None:
            customer_info[key] = ""

    #同意ページ
    check = driver.find_element_by_xpath("//input[@id='consent']")
    check.click()
    time.sleep(1)
    entry = driver.find_element_by_css_selector(".btn.btn-primary")
    entry.click()

    time.sleep(2)
    #お客様情報　要素取得
    elements_info = getElementsInfo(driver)

    #お客様情報入力#
    insertCustomerInfo(elements_info, customer_info, driver)
    time.sleep(80)
    confire_btn = driver.find_element_by_id("confire_btn")
    confire_btn.click()

    # /apply/inputページ#
    time.sleep(2)
    apply = driver.find_element_by_css_selector(".btn.btn-primary")
    apply.click()

    # /apply/identifyページ#
    time.sleep(2)
    main_identity = driver.find_element_by_xpath("//input[@name='main_identity'][@value='1']")
    main_identity.click()
    image1 = driver.find_element_by_name("image1")
    image2 = driver.find_element_by_name("image2")
    image1.send_keys(image_file)
    image2.send_keys(image_file)
    image4 = driver.find_element_by_name("image4")
    image4.send_keys(image_file)
    complete_btn = driver.find_element_by_id("complete_btn")
    complete_btn.click()

    #次の顧客登録
    time.sleep(2)
    start_column = start_column + 1
    driver.get(url)



