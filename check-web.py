from selenium import webdriver
## BY: 也就是依照條件尋找元素中XPATH、CLASS NAME、ID、CSS選擇器等都會用到的Library
from selenium.webdriver.common.by import By
## keys: 鍵盤相關的Library
from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.chrome.options import  Options
import time
import datetime as dt
import pandas as pd
import requests
from PIL import Image



test_line_token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI"
vhwc_line_token = "HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz"




def crop_image(image_path, crop_length):
    # 檢查圖片格式
    if not image_path.lower().endswith(".png"):
        raise ValueError("Image format must be PNG")

    # 讀取圖片
    img = Image.open(image_path)

    # 計算圖片的寬度和高度
    width, height = img.size

    # 計算切片的數量
    num_crops = (height // crop_length) + 1

    filepaths = []
    # 進行切片
    for i in range(num_crops):
        # 計算切片的左上角座標
        x = 0
        y = i * crop_length

        # 計算切片的右下角座標
        w = width
        h = y + crop_length

        # 進行切片
        cropped_img = img.crop((x, y, w, h))

        # 儲存切片
        filepath = f"{image_path[:-4]}_{i + 1}.png"
        cropped_img.save(filepath)
        filepaths.append(filepath)
        
    return filepaths


def send_to_line_notify_bot(msg, line_notify_token, photo_opened=None):
    url = "https://notify-api.line.me/api/notify"
    headers = {"Authorization": f"Bearer {line_notify_token}"}
    data = {"message":msg}
    image_file = {'imageFile': photo_opened}
    r = requests.post(url=url,data=data,headers=headers,files=image_file)

def check_oe(url,account,pwd):
    #

    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    #防止chrome自動?閉
    options.add_experimental_option(name="detach", value=True)
    #chrome 的無界面模式, 此模式才可以截長圖
    options.add_argument("headless")

    # 產生截圖檔名
    # 把172.20.1.12中的20或19取出,對應到vhwc或vhcy.
    hospital = {
        '19' : "vhcy",
        '20' : "vhwc"
    }
    url_content = url.split("/")
    ip_2 = url_content[2].split(".")[1]
    
    # 檔名結尾為日期時間附加.
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

    # 回傳網頁載入後的長寬
    width = driver.execute_script("return document.documentElement.scrollWidth")
    height = driver.execute_script("return document.documentElement.scrollHeight")
     
    driver.set_window_size(width, height) 
    time.sleep(1) 
    
    save_path = f"d:\mis\{png_filename}"
    driver.get_screenshot_as_file(save_path)
    
    #line notify 傳送圖片可能有限制, 過長會壓縮. 如果超過2500, 就截切圖片. 
    if height > 2040:
        captured_images = crop_image(image_path=save_path,crop_length=2040)
    else :
        captured_images = [save_path,]    
    

    # 截圖完成, 找錯誤log

    # 找出綱頁中的table, 轉為dataframe, 再將有"失敗"字串的資料取出.
    report_element = driver.find_element(By.CLASS_NAME,"tableIn")
    report_html = report_element.get_attribute('outerHTML')
    report_df = pd.read_html(report_html)[0]
    report_fail_list = report_df[report_df['執行狀態'].str.contains("失敗")]

    driver.close()


    #整理失敗的資料, 轉成要發送的訊息
    title_msg = f"{hospital[ip_2]} {url_content[3]} {now.strftime('%Y%m%d %H:%M:%S')}\n"
    if report_fail_list.empty:
        msg = "🟢 Pass"
    else:
        msg = f"🚨 Fail: 總共{report_fail_list.shape[0]}個\n"

        for r in range(report_fail_list.shape[0]):
            msg += f"ID: {report_fail_list.iloc[r,0]}\n\
                    說明: {report_fail_list.iloc[r,5]}\n\
                    ---------\n"
            
    send_msg = title_msg + msg

    # 回傳要傳line的訊息和截圖儲存路徑(可能有切圖)
    return {'msg' : send_msg,
            'filepath' : captured_images,
            #'df' : report_fail_list
            }



def check_showjob (url):
    
    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    #防止chrome自動?閉
    options.add_experimental_option(name="detach", value=True)
    #chrome 的無界面模式, 此模式才可以截長圖
    #options.add_argument("headless")

    #產生截圖檔名
    # 把172.20.1.12中的20或19取出,對應到vhwc或vhcy.
    hospital = {
        '19' : "vhcy",
        '20' : "vhwc"
    }
    url_content = url.split("/")
    ip_2 = url_content[2].split(".")[1]
    
    # 檔名結尾為日期時間附加.
    now = dt.datetime.now()
    png_filename = f"{hospital[ip_2]}_showjob_{now.strftime('%Y%m%d%H%M%S')}.png"
    #name rule ex: vhwc_showjob_20240226123705.png

    driver = webdriver.Chrome(options=options)
    #witdth 1000, showjob截圖後長度長, 長度any, 載入網頁後會變.
    driver.set_window_size(width=1000,height=700)
    driver.get(url=url)

    button_run = driver.find_element(By.ID, "btnExec")
    button_run.click()

    #停長一點, 除非有寫等待載入完的code
    time.sleep(5)

    width = driver.execute_script("return document.documentElement.scrollWidth")
    height = driver.execute_script("return document.documentElement.scrollHeight")
     
    driver.set_window_size(width, height) 
    time.sleep(1) 
    
    save_path = f"d:\mis\{png_filename}"
    driver.get_screenshot_as_file(save_path)

    #line notify 傳送圖片可能有限制, 過長會壓縮. 如果超過2500, 就截切圖片. 
    if height > 2040:
        captured_images = crop_image(image_path=save_path,crop_length=2040)
    else :
        captured_images = [save_path,]    

    #截圖完成, 找錯誤log
    report_table = pd.read_html(driver.page_source)[0]
    new_head = report_table.iloc[2]
    report_table = report_table.drop(report_table.columns[:3],axis=0)
    report_table.columns = new_head
    report_fail_table = report_table[report_table['執行時間'].str.contains("失敗")]

    driver.close()
    

    #整理reprot
    title_msg = f"{hospital[ip_2]} showjob {now.strftime('%Y%m%d %H:%M:%S')}\n"
    if report_fail_table.empty:
        msg = "🟢 Pass"
    else:
        msg = f"🚨 Fail: 總共{report_fail_table.shape[0]}個\n"

        for r in range(report_fail_table.shape[0]):
            msg += f"程式代碼: {report_fail_table.iloc[r,0]}\n\
                    執行狀況: {report_fail_table.iloc[r,6]}\n\
                    ---------\n"
            

    send_msg = title_msg + msg


    return {'filepath' : captured_images,
            'msg' : send_msg}


### 檢查cpoe
report = check_oe(url="http://172.20.200.71/cpoe/m2/batch",account=73058,pwd="Q1220416")

send_to_line_notify_bot(msg=report['msg'], line_notify_token=test_line_token,photo_opened=None)
for i in report["filepath"]:
    msg = f"{report['filepath'].index(i)} / {report["filepath"].shape}"
    send_to_line_notify_bot(msg=msg,line_notify_token=test_line_token,photo_opened=open(i,"rb"))
    
### 檢查eror
report = check_oe(url="http://172.20.200.71/eroe/m2/batch",account=73058,pwd="Q1220416")

send_to_line_notify_bot(msg=report['msg'], line_notify_token=test_line_token,photo_opened=None)
for i in report["filepath"]:
    msg = f"{report['filepath'].index(i)} / {report["filepath"].shape}"
    send_to_line_notify_bot(msg=msg,line_notify_token=test_line_token,photo_opened=open(i,"rb"))
    

report = check_showjob(url = "http://172.20.200.41/NOPD/showjoblog.aspx")

send_to_line_notify_bot(msg=report['msg'], line_notify_token=test_line_token,photo_opened=None)
for i in report["filepath"]:
    msg = f"{report['filepath'].index(i)} / {report["filepath"].shape}"
    send_to_line_notify_bot(msg=msg,line_notify_token=test_line_token,photo_opened=open(i,"rb"))