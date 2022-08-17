from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
import openpyxl
import unicodedata
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
from def3 import getCustomerInfo
from def3 import convertCustomerInfo
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import glob

capabilities = DesiredCapabilities.CHROME.copy()
capabilities['acceptInsecureCerts'] = True
##########################################################################################################
######
#設定#
######
url = "http://entrylocal.examobile.jp:8080/"
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome("C:\\Users\\kyosuke.ushiyama\\AppData\\Local\\SeleniumBasic\\chromedriver.exe", desired_capabilities=capabilities)
#identify挿入画像
image_file = ("C:\\Users\\kyosuke.ushiyama\\python\\license.jpg")

def getFileName():
    idir = 'C:\\User\\kyosuke.ushiyama\\python'
    filetype = [("エクセル","*.xlsx")]
    file_path = tk.filedialog.askopenfilename(filetypes = filetype, initialdir = idir)
    input_box.delete(0,tk.END)
    input_box.insert(tk.END, file_path)
    print(file_path)

def outputCustomerInfo():
    file_path = input_box.get()
    if not file_path:
        print("no file")
        return
    
    tree = ttk.Treeview(root)
    # 列インデックスの作成
    tree["columns"] = (1,2,3,4,5,6,7,8,9,10,11,12,13,14)
    # 表スタイルの設定(headingsはツリー形式ではない、通常の表形式)
    tree["show"] = "headings"
    # 各列の設定(インデックス,オプション(今回は幅を指定))
    tree.column(1,width=100)
    tree.column(2,width=100)
    tree.column(3,width=75)
    tree.column(4,width=75)
    tree.column(5,width=100)
    tree.column(6,width=50)
    tree.column(7,width=100)
    tree.column(8,width=125)
    tree.column(9,width=100)
    tree.column(10,width=100)
    tree.column(11,width=75)
    tree.column(12,width=100)
    tree.column(13,width=110)
    tree.column(14,width=75)

    # 各列のヘッダー設定(インデックス,テキスト)
    tree.heading(1,text="氏名(カナ)")
    tree.heading(2,text="氏名")
    tree.heading(3,text="郵便番号")
    tree.heading(4,text="町名・番地")
    tree.heading(5,text="電話番号")
    tree.heading(6,text="性別")
    tree.heading(7,text="生年月日")
    tree.heading(8,text="メールアドレス")
    tree.heading(9,text="IMEI")
    tree.heading(10,text="カード番号")
    tree.heading(11,text="カード有効期間")
    tree.heading(12,text="カード名義人(カナ)")
    tree.heading(13,text="カード名義人生年月日")
    tree.heading(14,text="セキュリティコード")
    #エクセル読み取りファイル指定
    book = openpyxl.load_workbook(file_path, data_only=True)    
    sheet = book["お客様情報2"]
    num_of_customers = int(sheet["C1"].value)

    start_row = 6
    start_column = 3

    customer_info_list = []

    for i in range(num_of_customers):
        #excelから顧客情報を読み取る
        customer_info = getCustomerInfo(start_row, start_column, sheet)
        #未入力項目(None)を空白に変換
        for key in customer_info:
            if customer_info[key] is None:
                customer_info[key] = ""

        tree.insert("","end",values=(customer_info["kana"],\
                                    customer_info["name"],\
                                    customer_info["zipcode"],\
                                    customer_info["s_b"],\
                                    customer_info["telno"],\
                                    customer_info["sex"],\
                                    customer_info["birth"],\
                                    customer_info["mail"],\
                                    customer_info["imei"],\
                                    customer_info["card_no"],\
                                    customer_info["card_expire"],\
                                    customer_info["card_kana"],\
                                    customer_info["card_birth"],\
                                    customer_info["security_cd"]
                                    )
                    )
        customer_info_list.append(customer_info)
        start_column = start_column + 1
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    scrollbar.place(x=10+200+2, y=200, height=200+20)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.place(x=10, y=200)

def register():
    file_path = input_box.get()
    if not file_path:    
        return
    
    driver.get(url)
    for i in range(num_of_customers):
    
        #excelから顧客情報を読み取る
        converted_customer_info = convertCustomerInfo(customer_info)
        #customer_info = getCustomerInfo(start_row, start_column, sheet)
        #未入力項目(None)を空白に変換
        for key in converted_customer_info:
            if converted_customer_info[key] is None:
                converted_customer_info[key] = ""

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
        insertCustomerInfo(elements_info, converted_customer_info, driver)
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
    

if __name__ == "__main__":
    root = tk.Tk()
    root.title("自動顧客登録")
    root.geometry("1300x600")
    frame = tk.Frame(root,width=1200, height=600)
    frame.pack()
    
    #ラベルの作成
    input_label = tk.Label(frame, text="顧客一覧のファイルを選んでください")
    input_label.place(relx=0.45, y=70)
    
    #入力欄の作成
    input_box = tk.Entry(frame, width=40)
    input_box.place(relx=0.4, y=100)

    #ボタンの作成
    button_search = tk.Button(frame, text="参照",command=getFileName)
    button_search.place(relx=0.63, y=100)

    button_output = tk.Button(frame, text="顧客一覧表示", command=lambda:outputCustomerInfo())
    button_output.place(relx=0.5, y=150)

    button_register = tk.Button(frame, text="顧客登録開始", command=register)
    button_register.place(relx=0.5, y=500)

    root.mainloop()

