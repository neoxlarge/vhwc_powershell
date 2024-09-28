<#
selenium 使用方法
https://github.com/adamdriscoll/selenium-powershell
https://www.zenrows.com/blog/selenium-powershell#interaction-automation

chromedrive.exe 下載位置
https://googlechromelabs.github.io/chrome-for-testing/#stable


#>


# 儲存截圖和網頁檔的路徑
$result_path = "\\172.20.5.185\mis\dutycheck_result"

# chromedrive 路徑, 此powershell預計會到遠端桌面主機上執行, 所以要確認遠端主機上的路徑.
# 預計是放在 d:\mis\vhwc_powershell\chromedriver.exe
$chromedriver_path = "d:\mis\vhwc_powershell\"

# selenium module path
# 遠端桌面主機未安裝selenium powershell 模組, 用匯入的方式載入模組.
Import-Module "d:\mis\vhwc_powershell\selenium\3.0.1\selenium.psd1"



$check_oe = @{
    'vhwc_cpoe' = @{
        'check_item'   = 'cpoe'
        'branch'       = "vhwc"
        'url'          = 'http://172.20.200.71/cpoe/m2/batch'
        'account'      = 'CC4F'
        'password'     = 'acervghtc'
        'capture_area' = 'end'
    };
    'vhcy_cpoe' = @{
        'check_item'   = 'cpoe'
        'branch'       = "vhcy"
        'url'          = 'http://172.19.200.71/cpoe/m2/batch'
        'account'      = 'CC4F'
        'password'     = 'acervghtc'
        'capture_area' = 'end'
    };

    'vhwc_eroe' = @{
        'check_item'   = 'eroe'
        'branch'       = "vhwc"
        'url'          = 'http://172.20.200.71/eroe/m2/batch'
        'account'      = 'CC4F'
        'password'     = 'acervghtc'
        'capture_area' = 'end'
    };
    'vhcy_eroe' = @{
        'check_item'   = 'eroe'
        'branch'       = "vhcy"
        'url'          = 'http://172.19.200.71/eroe/m2/batch'
        'account'      = 'CC4F'
        'password'     = 'acervghtc'
        'capture_area' = 'end'
    };
}  



function check-oe( $check_item, $branch, $url, $account, $password,  $capture_area) {

    # 開啟瀏覽器, headless 模式
    $driver = Start-SeChrome -WebDriverDirectory $chromedriver_path -headless
    # 開啟網址
    Enter-SeUrl -Url $url -Driver $Driver

    # 填入帳號密碼,按登入
    $driver.FindElementByXPath("//input[@name='login']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pass']").SendKeys($password)
    $driver.FindElementByXPath("//input[@name='m2Login_submit']").Click()
    Start-Sleep -Seconds 3

    # 調整視窗大小, 用以全螢幕截圖
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1080)

    # 從capture_area 判斷,網頁要捲動到的位置, 以sendkey home, end 實作
    $body = $driver.FindElementByTagName("body")
    $body.SendKeys([OpenQA.Selenium.Keys]::Control + [OpenQA.Selenium.Keys]::$capture_area) 
   
    Start-Sleep -Seconds 1
    
    # 儲存頁面和截圖
    $driver.PageSource | Out-File -FilePath "$($result_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile("$($result_path)\$($check_item)_$($branch)_$($date).png", "png")

    # 關閉瀏覽器 
    Stop-SeDriver -Driver $Driver

    # 回傳結果
    $result = @{"check_item" = $check_item; 
        "branch"             = $branch; 
        "date"               = $date; 
        "png_filepath"       = "$($result_path)\$($check_item)_$($branch)_$($date).png";
        "html_filepath"      = "$($result_path)\$($check_item)_$($branch)_$($date).html"
    }

    return $result        

}

$check_showjob = @{
    'vhwc_showjob' = @{
        'check_item'   = 'showjob'
        'branch'       = "vhwc"
        'url'          = 'http://172.20.200.41/NOPD/showjoblog.aspx'
        'capture_area' = 'home'
    }
    
    'vhcy_showjob' = @{
        'check_item'   = 'showjob'
        'branch'       = "vhcy"
        'url'          = 'http://172.19.200.41/NOPD/showjoblog.aspx'
        'capture_area' = 'home'
    }
        
}

function check-showjob ($check_item, $branch, $url) {

    # 開啟瀏覽器, headless 模式
    $driver = Start-SeChrome -WebDriverDirectory $chromedriver_path -headless
    # 開啟網址
    Enter-SeUrl -Url $url -Driver $Driver

    $driver.FindElementByXPath("//input[@id='btnExec']").Click()
    start-sleep -second 2

    # 調整視窗大小, 用以全螢幕截圖
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1080)
    
    # 從capture_area 判斷,網頁要捲動到的位置, 以sendkey home, end 實作
    $body = $driver.FindElementByTagName("body")
    $body.SendKeys([OpenQA.Selenium.Keys]::Control + [OpenQA.Selenium.Keys]::$capture_area) 
    start-sleep -second 1
    
    # 儲存頁面和截圖
    $driver.PageSource | Out-File -FilePath "$($result_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile( "$($result_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # 關閉瀏覽器 
    Stop-SeDriver -Driver $driver

    # 回傳結果
    $result = @{"check_item" = $check_item; 
        "branch"             = $branch; 
        "png_filepath"       = "$($result_path)\$($check_item)_$($branch)_showjob.png";
        "html_filepath"      = "$($result_path)\$($check_item)_$($branch)_showjob.html"
    }
    return $result        
}


$check_cyp2001 = @{
    'vhwc_cyp2001' = @{
        'check_item' = 'cyp2001'
        'branch'     = "wc"
        'url_login'  = 'http://172.19.1.21/medpt/medptlogin.php'
        'url_query'  = 'http://172.19.1.21/medpt/cyp2001.php'
        'account'    = '73058'
        'password'   = 'Q1220416'
    }

    'vhcy_cyp2001' = @{
        'check_item' = 'cyp2001'
        'branch'     = "cy"
        'url_login'  = 'http://172.19.1.21/medpt/medptlogin.php'
        'url_query'  = 'http://172.19.1.21/medpt/cyp2001.php'
        'account'    = '73058'
        'password'   = 'Q1220416'
    }
}

function check-cyp2001 ($check_item, $branch, $url_login, $url_query, $account, $password) {

    # 報表系統要先登入,才能查詢
    # 開啟瀏覽器, headless 模式
    $driver = Start-SeChrome -WebDriverDirectory $chromedriver_path -headless
    # 開啟網址
    Enter-SeUrl -Url $url_login -Driver $Driver

    # 填入帳號密碼,按登入
    $driver.FindElementByXPath("//input[@name='cn']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pw']").SendKeys($password)
    $driver.FindElementByXPath("//input[@type='submit']").Click()
    
    
    # 日期轉換為民國年/月/日, 例如113/09/26
    # 取得今天的日期
    $Today = Get-Date
    # 計算民國年
    $TaiwanYear = $Today.Year - 1911
    # 組合民國年、月、日
    $TaiwanDate = "{0:D3}/{1:D2}/{2:D2}" -f $TaiwanYear, $Today.Month, $Today.Day
   
    # 登入完, 用powershell發出requests 模擬post 查詢報表
    $body = @{
        'g_yyymmdd_s' = $TaiwanDate
        'from'        = $branch
    }
    
    # 報表查詢會花一些時間跑
    $result = Invoke-WebRequest -uri $url_query -Method POST -Body $body
    Start-Sleep -Seconds 2

    # 網頁存檔後再讀取
    $result.Content | Out-File -FilePath "$($result_path)\$($check_item)_$($branch)_$($date).html"
    Enter-SeUrl -Url "$($result_path)\$($check_item)_$($branch)_$($date).html" -Driver $driver

    # 取得視窗大小, 嘉義的比較長, 怕截到表格.
    $width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    $height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # 調整視窗大小, 用以全螢幕截圖
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(($width), ($height))
    # 截圖存檔
    $driver.GetScreenshot( ).SaveAsFile( "$($result_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # 關閉瀏覽器 
    Stop-SeDriver -Driver $driver

    # 回傳結果
    $result = @{"check_item" = $check_item;
        "branch"             = $branch;
        "png_filepath"       = "$($result_path)\$($check_item)_$($branch)_$($date).png";
        "html_filepath"      = "$($result_path)\$($check_item)_$($branch)_$($date).html"
    }
    return $result
}


# 取得日期, 以便命名檔案
$date = (get-date).ToString('yyyyMMddhhmm')


foreach ($key in $check_oe.keys) {
    check-oe -check_item $key -branch $check_oe[$key]['branch'] -url $check_oe[$key]['url'] -account $check_oe[$key]['account'] -password $check_oe[$key]['password'] -capture_area $check_oe[$key]['capture_area']
}

foreach ($key in $check_showjob.keys) {
    check-showjob -check_item $key -branch $check_showjob[$key]['branch'] -url $check_showjob[$key]['url']  -capture_area $check_showjob[$key]['capture_area']
}

foreach ($key in $check_cyp2001.keys) {
    check-cyp2001 -check_item $key -branch $check_cyp2001[$key]['branch'] -account $check_cyp2001[$key]['account'] -password $check_cyp2001[$key]['password'] -url_login $check_cyp2001[$key]['url_login'] -url_query $check_cyp2001[$key]['url_query'] 
}