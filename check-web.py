from selenium import webdriver
## BY: 也就是依照條件尋找元素中XPATH、CLASS NAME、ID、CSS選擇器等都會用到的Library
from selenium.webdriver.common.by import By
## keys: 鍵盤相關的Library
from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.chrome.options import  Options
import time
import datetime as dt
import pandas as pd





def check_oe(url,account,pwd):
    #

    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    #防止chrome自動?閉
    options.add_experimental_option(name="detach", value=True)
    #chrome 的無界面模式, 此模式才可以截長圖
    options.add_argument("headless")

    #產生截圖檔名
    hospital = {
        '19' : "vhcy",
        '20' : "vhwc"
    }
    url_content = url.split("/")
    
    ip_2 = url_content[2].split(".")[1]
    now = dt.datetime.now()
    png_filename = f"{hospital[ip_2]}_{url_content[3]}_{now.strftime('%Y%m%d%H%M%S')}.png"
    #name rule ex: vhwc_eroe_20240226123705.png

    driver = webdriver.Chrome(options=options)
    #witdth 1800, 截圖後長度比較剛好, 長度any, 載入網頁後會變.
    driver.set_window_size(width=1800,height=700)
    driver.get(url=url)

    input_account = driver.find_element(By.NAME,"login")
    input_pwd = driver.find_element(By.NAME,"pass")

    input_account.send_keys(account)
    input_pwd.send_keys(pwd)

    input_submit = driver.find_element(By.NAME,"m2Login_submit")
    input_submit.click()

    time.sleep(3)

    width = driver.execute_script("return document.documentElement.scrollWidth")
    height = driver.execute_script("return document.documentElement.scrollHeight")
     
    driver.set_window_size(width, height) 
    time.sleep(1) 
    
    save_path = f"d:\mis\{png_filename}"
    driver.get_screenshot_as_file(save_path)

    #截圖完成, 找錯誤log

    report_element = driver.find_element(By.CLASS_NAME,"tableIn")
    
    report_html = report_element.get_attribute('outerHTML')
    
    report_df = pd.read_html(report_html)[0]

    report_fail_list = report_df[report_df['執行狀態'].str.contains("失敗")]

    print(report_fail_list)
    driver.close()

    #return save_path
    return report_df

report = check_oe(url="http://172.20.200.71/cpoe/m2/batch",account=73058,pwd="Q1220416")
#check_oe(url="http://172.20.200.71/eroe/m2/batch",account=73058,pwd="Q1220416")

