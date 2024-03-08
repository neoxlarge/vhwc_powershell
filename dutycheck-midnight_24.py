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
from PIL import Image, ImageDraw, ImageFont
import argparse
import json
from io import StringIO


def add_watermark(image_path, text):
  """
  在 PNG 檔中加入文字浮水印

  Args:
    image_path: 要加入浮水印的 PNG 檔路徑
    text: 浮水印文字
    output_path: 加入浮水印後的 PNG 檔輸出路徑

  Returns:
    None
  """

  # 載入 PNG 檔
  image = Image.open(image_path)

  # 建立文字浮水印
  watermark = Image.new("RGBA", image.size, (0, 0, 0, 0))
  draw = ImageDraw.Draw(watermark)
  draw.text((0, 0), text, font=ImageFont.truetype("arial.ttf", 80), fill=(255, 25, 255, 128))

  # 將文字浮水印加入 PNG 檔
  image.paste(watermark, (0, 0), watermark)

  # 儲存加入浮水印後的 PNG 檔
  image.save(image_path)


def crop_image(image_path, crop_length):
    """把圖檔依長度切割, 存檔後回傳檔名路徑"""
    
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

   

def check_oe(url,account,pwd):
    """ 
    ### 檢查cpoe和eroe網頁 
    * 依傳入的綱址, 折出灣橋和嘉義, cpoe和eroe
    * 長螢幕截圖, 過長的圖, line 會壓縮, 造成糊掉, 圖檔會依長度切割
    * 會檢查表格中內格有出現"失敗"字串, 會傳訊息提醒
    """
    url_content = url.split("/")
    #url sample = http://172.20.200.71/cpoe/m2/batch

    branch_ipcode = {
        '19' : "vhcy",
        '20' : "vhwc" 
    }

    report = {
        "date" : dt.datetime.now().strftime('%Y%m%d'),
        "time" : dt.datetime.now().strftime('%H:%M:%S'),
        "taiwan_yyymmdd" : f"{dt.datetime.now().year - 1911}/{dt.datetime.now():%m}/{dt.datetime.now():%d}",
        "url" : url,
        "url_connected" : False,
        "branch" : branch_ipcode[url_content[2].split(".")[1]],
        "checkitem" : url_content[3],
        "png_foldername" : png_foldername,
        "png_filename" : None,
        "png_filepath" : None,
        "crop_images" : None,
        "fail_list" : None,
        "message" : None,
        "result" : False
    }
    
    # 產生截圖檔名, name rule ex: vhwc_eroe_20240226123705.png
    report['png_filename'] = f"{report['branch']}_{report['checkitem']}_{report['date']}{report['time'].replace(':','')}.png"
    report['png_filepath'] = f"{report['png_foldername']}{report['png_filename']}"

    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    #防止chrome自動關閉
    options.add_experimental_option(name="detach", value=True)
    #禁用devtools日誌
    options.add_experimental_option('excludeSwitches', ['disable-logging'])
    #chrome 的無界面模式, 此模式才可以截長圖
    options.add_argument("headless")
    #指定webdriver路徑
    service = webdriver.ChromeService(driver_path)
    #開啟chrome
    driver = webdriver.Chrome(options=options,service=service)
    #width 1800, 截圖後長度比較剛好, 長度any, 載入網頁後會變.
    driver.set_window_size(width=1800,height=700)
    driver.implicitly_wait(10)

    print(f"檢查 {report['url']}",end="")
    
    #檢查url是否可正常連線
    try:
        driver.get(url=url)
        report['url_connected'] = True
    except (WebDriverException, TimeoutException) as e:
        driver.close()
        report['url_connected'] = False
        report['message'] = "{url} 連線失敗"
    except Exception as e:
        driver.close()
        report['url_connected'] = False
        report['message'] = "不明原因失敗"
    
    print(f"檢查 {report['url_connected']}")        
        
    if report['url_connected']:    
        input_account = driver.find_element(By.NAME,"login")
        input_pwd = driver.find_element(By.NAME,"pass")

        input_account.send_keys(account)
        input_pwd.send_keys(pwd)

        input_submit = driver.find_element(By.NAME,"m2Login_submit")
        input_submit.click()
        #time.sleep(3)

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
        
        for img in report['crop_images']:
            add_watermark(img, f"{report['branch']} {report['checkitem']} {report['date']}")
            
        # 截圖完成, 找錯誤log

        # 找出綱頁中的table, 轉為dataframe, 再將有"失敗"字串的資料取出.
        report_element = driver.find_element(By.CLASS_NAME,"tableIn")
        report_html = report_element.get_attribute('outerHTML')
        report_df = pd.read_html(StringIO(report_html))[0]
        report['fail_list'] = report_df[report_df['執行狀態'].str.contains("失敗")]

        driver.close()
        
        #整理失敗的資料, 轉成要發送的訊息
                
        if report['fail_list'].empty:
            report['result'] = True
            report['message'] = "Pass"
            report['fail_list'] = report['fail_list'].to_string() 
            
            
        else:
            report['result'] = False
            msg = f"總共{report['fail_list'].shape[0]}個\n"

            #for r in range(report_fail_list.shape[0]):
            for r in range(report['fail_list'].shape[0]):
                msg += f"ID: {report['fail_list'].iloc[r,0]}\n說明: {report['fail_list'].iloc[r,5]}\n---------\n"
                
            title_msg = f"{report['branch']} {report['checkitem']}\n ==={report['date']} {report['time']}===\n"            
            report['message'] = title_msg + msg
            report['fail_list'] = report['fail_list'].to_string() 
                
    # 回傳要傳line的訊息和截圖儲存路徑(可能有切圖)
    return report



def check_showjob (url):
    """ 
    ### 檢查showjob網頁 
        * 長螢幕截圖, 過長的圖, line 會壓縮, 造成糊掉, 圖檔會依長度切割
        * 會檢查表格中內格有出現"失敗"字串, 會傳訊息提醒
    """
    url_content = url.split("/")
    #url sample = http://172.20.200.41/NOPD/showjoblog.aspx

    branch_ipcode = {
        '19' : "vhcy",
        '20' : "vhwc" 
    }

    report = {
        "date" : dt.datetime.now().strftime('%Y%m%d'),
        "time" : dt.datetime.now().strftime('%H:%M:%S'),
        "taiwan_yyymmdd" : f"{dt.datetime.now().year - 1911}/{dt.datetime.now():%m}/{dt.datetime.now():%d}",
        "url" : url,
        "url_connected" : False,
        "branch" : branch_ipcode[url_content[2].split(".")[1]],
        "checkitem" : "showjob",
        "png_foldername" : png_foldername,
        "png_filename" : None,
        "png_filepath" : None,
        "crop_images" : None,
        "fail_list" : None,
        "message" : None,
        "result" : False
    }
    
    # 產生截圖檔名, name rule ex: vhwc_showjob_20240226123705.png
    report['png_filename'] = f"{report['branch']}_{report['checkitem']}_{report['date']}{report['time'].replace(':','')}.png"
    report['png_filepath'] = f"{report['png_foldername']}{report['png_filename']}"
        
    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    # 防止chrome自動關閉
    options.add_experimental_option(name="detach", value=True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # chrome 的無界面模式, 此模式才可以截長圖
    options.add_argument("headless")
    #指定webdriver路徑
    service = webdriver.ChromeService(driver_path)
    #開啟chrome
    driver = webdriver.Chrome(options=options,service=service)
    #width 1000, showjob截圖後長度長, 長度any, 載入網頁後會變.
    driver.set_window_size(width=1000,height=700)
    driver.implicitly_wait(10)

    print(f"檢查 {report['url']}",end="")

    #檢查url是否可正常連線
    try:
        driver.get(url=url)
        report['url_connected'] = True
    except (WebDriverException, TimeoutException) as e:
        driver.close()
        report['url_connected'] = False
        report['message'] = "{url} 連線失敗"
    except Exception as e:
        driver.close()
        report['url_connected'] = False
        report['message'] = "不明原因失敗"

    print(f"檢查 {report['url_connected']}")

    if report['url_connected']:

        button_run = driver.find_element(By.ID, "btnExec")
        button_run.click()
        #停長一點, 除非有寫等待載入完的code
        #time.sleep(5)

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
        
        for img in report['crop_images']:
            add_watermark(img, f"{report['branch']} {report['checkitem']} {report['date']}")
     

        #截圖完成, 找錯誤log
        report_table = pd.read_html(StringIO(driver.page_source))[0]
        new_head = report_table.iloc[2]
        report_table = report_table.drop(report_table.columns[:3],axis=0)
        report_table.columns = new_head
        report['fail_list'] = report_table[report_table['執行時間'].str.contains("失敗")]

        driver.close()
    
        #整理report
        if report['fail_list'].empty:
            report['result'] = True
            report['fail_list'] = report['fail_list'].to_string() 
        else:
            report['result'] = False
            msg = f"總共{report['fail_list'].shape[0]}個\n"

            for r in range(report['fail_list'].shape[0]) :
                msg += f"程式代碼: {report['fail_list'].iloc[r,0]}\n執行狀況: {report['fail_list'].iloc[r,6]}\n---------\n"
            
            title_msg = f"{report['branch']} {report['checkitem']}\n ==={report['date']} {report['time']}===\n"
            report['message'] = title_msg + msg
            report['fail_list'] = report['fail_list'].to_string() 

    return report


def check_cyp2001(branch,account,pwd):
    """ 
    ### 檢查外掛程式中處方log統計網頁 
    * 外掛不好用selenium取得網頁內容, 配合用requests取得綱頁, 
    先存成html檔, 再用selenium開啟並截圖     
    """
    report = {
        "date" : dt.datetime.now().strftime('%Y%m%d'),
        "time" : dt.datetime.now().strftime('%H:%M:%S'),
        "taiwan_yyymmdd" : f"{dt.datetime.now().year - 1911}/{dt.datetime.now():%m}/{dt.datetime.now():%d}",
        "url" : "http://172.19.1.21/medpt/medptlogin.php",
        "url_connected" : False,
        "branch" : branch,
        "checkitem" : "Prescription_log",
        "png_foldername" : png_foldername,
        "png_filename" : None,
        "png_filepath" : None,
        "crop_images" : None,
        "fail_list" : None,
        "message" : None,
        "result" : False

    }
    
        
    #檢查外掛報表
    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    #防止chrome自動關閉
    options.add_experimental_option(name="detach", value=True)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    #chrome 的無界面模式, 此模式才可以截長圖
    options.add_argument("headless")
    #指定webdriver路徑
    service = webdriver.ChromeService(driver_path)
    #開啟chrome
    driver = webdriver.Chrome(options=options,service=service)
    #width 600, 外掛表格比較窄, 長度any, 載入網頁後會變.
    driver.set_window_size(width=400,height=600)
    
    print(f"登入 {report['url']} branch: {report['branch']}", end="")
    
    #檢查url是否可正常連線
    try:
        driver.get(url=report['url'])
        report['url_connected'] = True
    except (WebDriverException, TimeoutException) as e:
        driver.close()
        report['url_connected'] = False
        report['message'] = f"{report['url']} 連線失敗"
    except Exception as e:
        driver.close()
        report['url_connected'] = False
        report['message'] = "不明原因失敗"

    print(f"檢查 {report['url_connected']}")
    
    if report['url_connected']:
        #第二層, 連去cyp2001.php, 取得cyp2001表格內容.
        loginname = driver.find_element(By.NAME,"cn")
        loginpwd = driver.find_element(By.NAME,"pw")
        loginname.send_keys(account)
        loginpwd.send_keys(pwd)
        loginpwd.send_keys(Keys.RETURN) #直接按enter送出
    
        time.sleep(1)
        
        report['png_filename'] = f"vh{report['branch']}_{report['checkitem']}_{report['date']}{report['time'].replace(':', '')}.png"
        report['png_filepath'] = f"{report['png_foldername']}{report['png_filename']}"
        

        url = "http://172.19.1.21/medpt/cyp2001.php"
        data = {'g_yyymmdd_s': report['taiwan_yyymmdd'],'from': {report['branch']},}

        save_html_path = f"{report['png_filepath'].replace('.png','.html')}"
        
        print(f"檢查 {url} branch: {report['branch']}", end="")
        
        try:
            response = requests.post(url=url,data=data)
            response.raise_for_status()
        except requests.exceptions.HTTPError as err:        #有問題可能是cyp2001.php有問題.
            report['url_connected'] = False
            report['message'] = f"vh{report['branch']} {report['checkitem']} \n ==={report['date']} {report['time']}===\n {url} 連線失敗"
            
        print(f"檢查 {response.status_code}")    
            
        if response.status_code == 200: #code 200 表示網頁正確取得, 寫入html檔.
            with open(save_html_path, 'wb') as f:
                f.write(response.content)

            driver.get(save_html_path)

            width = driver.execute_script("return document.documentElement.scrollWidth")
            height = driver.execute_script("return document.documentElement.scrollHeight")
            driver.set_window_size(width, height) 
            
            time.sleep(1)
            driver.save_screenshot(report['png_filepath'])

            report['result'] = True
            
        else:
            report['result'] = False
            report['message'] = f"vh{report['branch']} {report['checkitem']} \n ==={report['date']} {report['time']}===\n {url} 網頁回應不正確:\n {response.status_code}"

    driver.close()

    return report
    


### 檢查處方LOG統計
    #只有早上0點30分需要檢查這個
#now = dt.datetime.now()
#if now.hour <=1:    
#check_cyp2001(account=73058,pwd="Q1220416")    
    
def main(): 
    parser = argparse.ArgumentParser(description='傳入webdriver.exe路徑和圖片存檔資料夾')
    parser.add_argument('--driver_path', type=str, default='d:\\mis\\webdriver\\chromedriver.exe', help='webdriver.exe路徑',required=False)
    parser.add_argument('--png_foldername', type=str, default='d:\\mis\\', help='圖片存檔資料夾',required=False)
    args = parser.parse_args(['--driver_path','d:\\mis\\webdriver\\chromedriver.exe','--png_foldername','d:\\mis\\'])
    global png_foldername, driver_path
    png_foldername = args.png_foldername
    driver_path = args.driver_path
    
    ################################
    print("VHWC/VHCY 值班截圖")
    
    report_list = []

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
    
    for check in check_list:
        report = check_oe(url=check['url'], account=check['account'], pwd=check['pwd'])
        report_list.append(report)
                      
    #檢查嘉義和灣橋的所有showjob
    check_list = [{'url' : 'http://172.20.200.41/NOPD/showjoblog.aspx'},
                {'url' : 'http://172.19.200.41/NOPD/showjoblog.aspx'}] 
    
    for check in check_list:
        report = check_showjob(url=check['url'])
        report_list.append(report)

    check_list = ['wc','cy']
    for check in check_list:
        report = check_cyp2001(account=73058, pwd="Q1220416", branch=check)
        report_list.append(report)    
    
    #把report_list存檔到png_foldername資料夾,格式是json, 檔名是dutycheck.json
    json.dump(report_list, open(f"{png_foldername}dutycheck.json", "w"))
    #print(report_list)        
    



if __name__ == "__main__": 
    main()