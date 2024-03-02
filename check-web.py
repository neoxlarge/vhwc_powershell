from selenium import webdriver
## BY: 也就是依照條件尋找元素中XPATH、CLASS NAME、ID、CSS選擇器等都會用到的Library
from selenium.webdriver.common.by import By
## keys: 鍵盤相關的Library
from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.chrome.options import  Options
from selenium.common.exceptions import WebDriverException, TimeoutException

import time
import datetime as dt
import pandas as pd
import requests
from PIL import Image



test_line_token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI"
vhwc_line_token = "HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz"
vhwc_line_token = test_line_token



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
    url_content = url.split("/")
    #url sample = http://172.20.200.71/cpoe/m2/batch

    branch_ipcode = {
        '19' : "vhcy",
        '20' : "vhwc" 
    }

    report = {
        "date" : dt.datetime.now().strftime('%Y%m%d'),
        "time" : dt.datetime.now().strftime('%H:%M:%S'),
        "url" : url,
        "url_connected" : False,
        "branch" : branch_ipcode[url_content[2].split(".")[1]],
        "oe" : url_content[3],
        "png_foldername" : "d:\\mis\\",
        "png_filename" : None,
        'png_filepath' : None,
        'crop_images' : None,
        'fail_list' : None,
        "message" : None
    }
    
    # 產生截圖檔名, name rule ex: vhwc_eroe_20240226123705.png
    report['png_filename'] = f"{report['branch']}_{report['oe']}_{report['date']}{report['time'].replace(':','')}.png"
    report['png_filepath'] = f"{report['png_foldername']}{report['png_filename']}"

    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    #防止chrome自動關閉
    options.add_experimental_option(name="detach", value=True)
    #chrome 的無界面模式, 此模式才可以截長圖
    options.add_argument("headless")

    driver = webdriver.Chrome(options=options)
    #width 1800, 截圖後長度比較剛好, 長度any, 載入網頁後會變.
    driver.set_window_size(width=1800,height=700)
    #檢查url是否可正常連線
    try:
        driver.get(url=url)
        report['url_connected'] = True
        
        
    except (WebDriverException, TimeoutException) as e:
        driver.close()
        msg = f"🚨 Fail: {url} 連線失敗"
        #report['url_connected'] = False
        
    if report['url_connected']:    
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
        
        driver.get_screenshot_as_file(report['png_filepath'])
        
        
        #line notify 傳送圖片可能有限制, 過長會壓縮. 如果超過2500, 就截切圖片. 
        if height > 2040:
            report['crop_images'] = crop_image(image_path=report['png_filepath'], crop_length=2040) 
        else:
            report['crop_images'] = [report['png_filepath'],]
        
        # 截圖完成, 找錯誤log

        # 找出綱頁中的table, 轉為dataframe, 再將有"失敗"字串的資料取出.
        report_element = driver.find_element(By.CLASS_NAME,"tableIn")
        report_html = report_element.get_attribute('outerHTML')
        report_df = pd.read_html(report_html)[0]
        report['fail_list'] = report_df[report_df['執行狀態'].str.contains("失敗")]

        driver.close()
        
        #整理失敗的資料, 轉成要發送的訊息
                
        if report['fail_list'].empty:
            msg = "🟢 Pass"
        else:
            msg = f"🚨 Fail: 總共{report['fail_list'].shape[0]}個\n"

            #for r in range(report_fail_list.shape[0]):
            for r in range(report['fail_list'].shape[0]):
                msg += f"ID: {report['fail_list'].iloc[r,0]}\n說明: {report['fail_list'].iloc[r,5]}\n---------\n"
                
    title_msg = f"{report['branch']} {report['oe']}\n ==={report['date']} {report['time']}===\n"            
    report['message'] = title_msg + msg
        
    # 回傳要傳line的訊息和截圖儲存路徑(可能有切圖)
    return report



def check_showjob (url):
    
    url_content = url.split("/")
    #url sample = http://172.20.200.41/NOPD/showjoblog.aspx

    branch_ipcode = {
        '19' : "vhcy",
        '20' : "vhwc" 
    }

    report = {
        "date" : dt.datetime.now().strftime('%Y%m%d'),
        "time" : dt.datetime.now().strftime('%H:%M:%S'),
        "url" : url,
        "url_connected" : False,
        "branch" : branch_ipcode[url_content[2].split(".")[1]],
        "item" : "showjob",
        "png_foldername" : "d:\\mis\\",
        "png_filename" : None,
        'png_filepath' : None,
        'crop_images' : None,
        'fail_list' : None,
        "message" : None
    }
    
    # 產生截圖檔名, name rule ex: vhwc_showjob_20240226123705.png
    report['png_filename'] = f"{report['branch']}_{report['item']}_{report['date']}{report['time'].replace(':','')}.png"
    report['png_filepath'] = f"{report['png_foldername']}{report['png_filename']}"
    
    
    
    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    #防止chrome自動關閉
    options.add_experimental_option(name="detach", value=True)
    #chrome 的無界面模式, 此模式才可以截長圖
    options.add_argument("headless")

    driver = webdriver.Chrome(options=options)
    #witdth 1000, showjob截圖後長度長, 長度any, 載入網頁後會變.
    driver.set_window_size(width=1000,height=700)
    try:
        driver.get(url=url)
        report['url_connected'] = True
        
    except (WebDriverException,TimeoutException) as e:
        driver.close()
        msg = f"🚨 Fail: {url} 連線失敗"    

    if report['url_connected']:

        button_run = driver.find_element(By.ID, "btnExec")
        button_run.click()

        #停長一點, 除非有寫等待載入完的code
        time.sleep(5)

        width = driver.execute_script("return document.documentElement.scrollWidth")
        height = driver.execute_script("return document.documentElement.scrollHeight")
        
        driver.set_window_size(width, height) 
        time.sleep(1) 
        
        driver.get_screenshot_as_file(report['png_filepath'])

        #line notify 傳送圖片可能有限制, 過長會壓縮. 如果超過2500, 就截切圖片. 
        if height > 2040:
            report['crop_images'] = crop_image(image_path=report['png_filepath'],crop_length=2040)
        else :
            report['crop_images'] = [report['png_filepath'],]    

        #截圖完成, 找錯誤log
        report_table = pd.read_html(driver.page_source)[0]
        new_head = report_table.iloc[2]
        report_table = report_table.drop(report_table.columns[:3],axis=0)
        report_table.columns = new_head
        report['fail_list'] = report_table[report_table['執行時間'].str.contains("失敗")]

        driver.close()
    

        #整理report
        if report['fail_list'].empty:
            msg = "🟢 Pass"
        else:
            msg = f"🚨 Fail: 總共{report['fail_list'].shape[0]}個\n"

            for r in range(report['fail_list'].shape[0]) :
                msg += f"程式代碼: {report['fail_list'].iloc[r,0]}\n執行狀況: {report['fail_list'].iloc[r,6]}\n---------\n"
            
    title_msg = f"{report['branch']} showjob\n ==={report['date']} {report['time']}===\n"
    report['message'] = title_msg + msg

    return report


def check_cyp2001(account,pwd):
    
    report = {
        "date" : dt.datetime.now().strftime('%Y%m%d'),
        "time" : dt.datetime.now().strftime('%H:%M:%S'),
        'taiwan_yyymmdd' : f"{dt.datetime.now().year - 1911}/{dt.datetime.now():%m}/{dt.datetime.now():%d}",
        "url" : "http://172.19.1.21/medpt/medptlogin.php",
        "url_connected" : False,
        "branch" : ['wc','cy'],
        "item" : "Prescription_log",
        "png_foldername" : "d:\\mis\\",
        #"html_filename" : None,
        #"png_filename" : None,
        #'png_filepath' : None,
        'crop_images' : None,
        'fail_list' : None,
        "message" : None
    }
    
        
    #檢查外掛報表
    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    #防止chrome自動關閉
    options.add_experimental_option(name="detach", value=True)
    #chrome 的無界面模式, 此模式才可以截長圖
    options.add_argument("headless")

    #先開chrome登入外掛系統.
    driver = webdriver.Chrome(options=options)
    #witdth 600, 外掛表格比較窄, 長度any, 載入網頁後會變.
    driver.set_window_size(width=400,height=600)
    
    try:
        driver.get(url=report['url'])
        report['url_connected'] = True
    except (WebDriverException, TimeoutException) as e:
        driver.close()
        msg = f"外掛糸統\n🚨 Fail: {url} 連線失敗"    
        send_to_line_notify_bot(msg=msg, line_notify_token=vhwc_line_token, photo_opened=None)


    if report['url_connected']:
        
        loginname = driver.find_element(By.NAME,"cn")
        loginpwd = driver.find_element(By.NAME,"pw")
        loginok = driver.find_element(By.CSS_SELECTOR,'input[value="確定"]')

        loginname.send_keys(account)
        loginpwd.send_keys(pwd)
        loginpwd.send_keys(Keys.RETURN)
    
        time.sleep(1)
        for b in report['branch'] :
    
            now = dt.datetime.now()
            path_title = f"{report['png_foldername']}vh{b}_{report['item']}_{report['taiwan_yyymmdd'].replace('/','')}"

            url = "http://172.19.1.21/medpt/cyp2001.php"
            data = {'g_yyymmdd_s': report['taiwan_yyymmdd'],'from': b,}

            save_html_path = f"{path_title}.html"
            save_img_path = f"{path_title}.png"

            try:
                response = requests.post(url=url,data=data)
                response.raise_for_status()
            except requests.exceptions.HTTPError as err:
                msg = f"vh{b} 處方LOG統計 \n ==={now.strftime('%Y%m%d %H:%M:%S')}===\n🚨 Fail: {url} 連線失敗"
                send_to_line_notify_bot(msg=msg, line_notify_token=vhwc_line_token, photo_opened=None)


            if response.status_code == 200:
                with open(save_html_path, 'wb') as f:
                    f.write(response.content)

                driver.get(save_html_path)

                width = driver.execute_script("return document.documentElement.scrollWidth")
                height = driver.execute_script("return document.documentElement.scrollHeight")
                driver.set_window_size(width, height) 
                
                time.sleep(1)
                driver.save_screenshot(save_img_path)

                msg = f"vh{b} 處方LOG統計 \n ==={now.strftime('%Y%m%d %H:%M:%S')}==="
                send_to_line_notify_bot(msg=msg, line_notify_token=vhwc_line_token, photo_opened=open(save_img_path, "rb"))

    driver.close()
    

def check_all_oe(check_list):
    for check in check_list:
        report = check_oe(url=check['url'], account=check['account'],pwd=check['pwd'])
        
        send_to_line_notify_bot(msg=report['message'], line_notify_token=vhwc_line_token,photo_opened=None)
        if report['crop_images']:
            for i in report["crop_images"]:
                msg = f"{report['branch']} {report['oe']} {report['crop_images'].index(i) + 1} / {len(report['crop_images'])}"
                send_to_line_notify_bot(msg=msg, line_notify_token=vhwc_line_token, photo_opened=open(i, "rb"))
                
                
def check_all_showjob(check_list):
    for check in check_list:
        report = check_showjob(url=check['url'])

        send_to_line_notify_bot(msg=report['message'], line_notify_token=vhwc_line_token, photo_opened=None)
        if report['crop_images']:
            for i in report["crop_images"]:
                msg = f"{report['branch']} showjob {report['crop_images'].index(i) + 1} / {len(report['crop_images'])}"
                send_to_line_notify_bot(msg=msg, line_notify_token=vhwc_line_token, photo_opened=open(i, "rb"))

#檢查嘉義和灣橋的所有oe
check_list = [{'url':"http://172.20.200.71/cpoe/m2/batch",
               'account' :  'CC4F',
               'pwd' : 'acervghtc'},
               {'url':"http://172.20.200.71/eroe/m2/batch",
               'account' :  'CC4F',
               'pwd' : 'acervghtc'},
               {'url':"http://172.19.200.71/cpoe/m2/batch",
               'account' :  'CC4F',
               'pwd' : 'acervghtc'},
               {'url':"http://172.19.200.71/eroe/m2/batch",
               'account' :  'CC4F',
               'pwd' : 'acervghtc'}
               ]

check_all_oe(check_list)

#檢查嘉義和灣橋的所有showjob
check_list = [{'url' : 'http://172.20.200.41/NOPD/showjoblog.aspx'},
              {'url' : 'http://172.19.200.41/NOPD/showjoblog.aspx'}] 
               
check_all_showjob(check_list)




### 檢查處方LOG統計
    #只有早上0點30分需要檢查這個
now = dt.datetime.now()
if now.hour <=1:    
    check_cyp2001(account=73058,pwd="Q1220416")    
    
