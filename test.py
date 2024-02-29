from selenium import webdriver

url="http://172.19.1.21/medpt/cyp2001.php"

#檢查外掛報表

options = webdriver.ChromeOptions()
#防止chrome自動?閉
options.add_experimental_option(name="detach", value=True)
#chrome 的無界面模式, 此模式才可以截長圖
#options.add_argument("headless")
branch = {'vhwc' : '灣橋',
        'vhcy' : '嘉義' }

 
#png_filename = f"{hospital[ip_2]}_{url_content[3]}_{now.strftime('%Y%m%d%H%M%S')}.png"
#name rule ex: vhwc_eroe_20240226123705.png


# 設定網址
url = 'http://172.19.1.21/medpt/cyp2001.php'

# 設定 POST 資料
data = {
    'g_yyymmdd_s': '113/02/29',
    'from': 'cy',
}





driver = webdriver.Chrome(options=options)
#witdth 1800, 截圖後長度比較剛好, 長度any, 載入網頁後會變.
driver.set_window_size(width=1800,height=700)
    
# 開啟網址
driver.get(full_url)
