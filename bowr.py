import os
import time

import pywinauto
from pywinauto.keyboard import send_keys
from selenium import webdriver
from selenium.webdriver.common.by import By
from userinfo import get_login_info

user_list = get_login_info().get_user_info()
# if not os.path.exists('info.txt.txt'):
#     os.mkfifo('info.txt.txt')

for item in user_list:
    username = item.get("username")
    if username in users:
        print("用户：{}已上传过".format(username))
        pass
    password = item.get("password")
    shop_id = item.get("shop_id")
    browser = webdriver.Chrome()
    browser.get('https://seller.octopia.com/login/')
    browser.find_element(by=By.ID, value="Login").send_keys(username)
    browser.find_element(by=By.ID, value="Password").send_keys(password)
    browser.find_element(by=By.ID, value="save").click()
    browser.get("https://seller.octopia.com/Product/Create/Massive/")
    browser.find_element(by=By.NAME, value="browseFile").click()

    # 使用pywinauto来选择文件
    app = pywinauto.Desktop()
    # 选择文件上传的窗口
    dlg = app["打开"]
    # 选择文件地址输入框，点击激活
    dlg["Toolbar3"].click()
    # 键盘输入上传文件的路径
    send_keys(r"D:\PythonProject\yz")
    # 键盘输入回车，打开该路径
    send_keys("{VK_RETURN}")
    # 选中文件名输入框，输入文件名
    dlg["文件名(&N):Edit"].type_keys("int_pdt_seller_import_%s.xlsm" % shop_id)
    # 点击打开
    dlg["打开(&O)"].click()
    browser.find_element(by=By.NAME, value="uploadFile").click()
    js = ""
    asd = browser.find_element(by=By.ID, value="AjaxResultPopinId").get_attribute("style")
    print(asd)
    file = open(file='info.txt', mode='w', encoding='utf-8')
    time.sleep(5)
    AjaxResultPopinIdStyle = browser.find_element(by=By.ID, value="AjaxResultPopinId").get_attribute("style")
    if "none" not in AjaxResultPopinIdStyle:
        error_info = browser.find_element(by=By.XPATH, value="//div[contains(@class,'modal-body')]").text
        if u"Une erreur est survenue lors du chargement du fichier. Veuillez réitérer l'opération" in error_info:
            print("上传失败！")

    else:
        pass
    file.write(username + "\n")
    browser.close()
