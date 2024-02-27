from selenium import webdriver
## BY: �]�N�O�̷ӱ���M�䤸����XPATH�BCLASS NAME�BID�BCSS��ܾ������|�Ψ쪺Library
from selenium.webdriver.common.by import By
## keys: ��L������Library
from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.chrome.options import  Options
import time
import datetime as dt
import pandas as pd





def check_oe(url,account,pwd):
    #

    # https://g.co/gemini/share/ada92acb29a0
    options = webdriver.ChromeOptions()
    #����chrome�۰�?��
    options.add_experimental_option(name="detach", value=True)
    #chrome ���L�ɭ��Ҧ�, ���Ҧ��~�i�H�I����
    options.add_argument("headless")

    #���ͺI���ɦW
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
    #witdth 1800, �I�ϫ���פ����n, ����any, ���J������|��.
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

    #�I�ϧ���, ����~log

    report_element = driver.find_element(By.CLASS_NAME,"tableIn")
    
    report_html = report_element.get_attribute('outerHTML')
    
    report_df = pd.read_html(report_html)[0]

    report_fail_list = report_df[report_df['���檬�A'].str.contains("����")]

    print(report_fail_list)
    driver.close()

    #return save_path
    return report_df

report = check_oe(url="http://172.20.200.71/cpoe/m2/batch",account=73058,pwd="Q1220416")
#check_oe(url="http://172.20.200.71/eroe/m2/batch",account=73058,pwd="Q1220416")

