<#
selenium 使用方法
https://github.com/adamdriscoll/selenium-powershell
https://www.zenrows.com/blog/selenium-powershell#interaction-automation




#>

Import-Module selenium

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
        'capture_area' = 'home'
    };
    'vhcy_eroe' = @{
        'check_item'   = 'eroe'
        'branch'       = "vhcy"
        'url'          = 'http://172.19.200.71/eroe/m2/batch'
        'account'      = 'CC4F'
        'password'     = 'acervghtc'
        'capture_area' = 'home'
    };
}  
}


function check-oe( $check_item, $branch, $url, $account, $password, $save_path ) {

    # 開啟瀏覽器, headless 模式
    $driver = Start-SeChrome -WebDriverDirectory ".\" -headless
    # 開啟網址
    Enter-SeUrl -Url $url -Driver $Driver

    # 填入帳號密碼,按登入
    $driver.FindElementByXPath("//input[@name='login']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pass']").SendKeys($password)
    $driver.FindElementByXPath("//input[@name='m2Login_submit']").Click()
    Start-Sleep -Seconds 3

    # 取得登入後的頁面的大小
    #$width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    #$height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # 調整視窗大小, 用以全螢幕截圖
    #$driver.Manage().Window.Size = New-Object System.Drawing.Size(($width + 120), ($height+200))
    driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1080)

    # 從capture_area 判斷,網頁要捲動到的位置, 以sendkey home, end 實作
    Send-SeKeys -Keys $check_oe[$key]['capture_area'] -Driver $Driver
    Start-Sleep -Seconds 2
    
    # 儲存頁面和截圖
    $driver.PageSource | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # 關閉瀏覽器 
    Stop-SeDriver -Driver $Driver

    # 回傳結果
    $result = @{"check_item" = $check_item; 
        "branch"             = $branch; 
        "date"               = $date; 
        "png_filepath"       = "$($save_path)\$($check_item)_$($branch)_$($date).png";
        "html_filepath"      = "$($save_path)\$($check_item)_$($branch)_$($date).html"
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

function check-showjob ($check_item, $branch, $url, $save_path) {

    # 開啟瀏覽器, headless 模式
    $driver = Start-SeChrome -WebDriverDirectory ".\" -headless
    # 開啟網址
    Enter-SeUrl -Url $url -Driver $Driver

    $driver.FindElementByXPath("//input[@id='btnExec']").Click()

    start-sleep -second 5

    # 取得登入後的頁面的大小
    #$width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    #$height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # 調整視窗大小, 用以全螢幕截圖
    #$driver.Manage().Window.Size = New-Object System.Drawing.Size(($width + 120), ($height+200))
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1080)
    
    # 從capture_area 判斷,網頁要捲動到的位置, 以sendkey home, end 實作
    Send-SeKeys -Keys $check_showjob[$key]['capture_area'] -Driver $Driver
    start-sleep -second 2
    
    # 儲存頁面和截圖
    $driver.PageSource | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # 關閉瀏覽器 
    Stop-SeDriver -Driver $driver

    # 回傳結果
    $result = @{"check_item" = $check_item; 
        "branch"             = $branch; 
        "png_filepath"       = "$($save_path)\$($check_item)_$($branch)_showjob.png";
        "html_filepath"      = "$($save_path)\$($check_item)_$($branch)_showjob.html"
    }
    return $result        
}


$check_cyp2001 = @{
    'vhwc_cyp2001' = @{
        'check_item' = 'cyp2001'
        'branch'     = "wc"
        'url_login'  = 'http://172.19.1.21/medpt/medptlogin.php'
        'url_query'  = 'http://172.19.1.21/medpt/medpt_cyp2001.php'
        'account'    = '73058'
        'password'   = 'Q1220416'
    }
}

function check-cyp2001 ($check_item, $branch, $url_login, $url_query, $account, $password, $save_path) {

    # 報表系統要先登入,才能查詢
    # 開啟瀏覽器, headless 模式
    $driver = Start-SeChrome -WebDriverDirectory ".\" -headless
    # 開啟網址
    Enter-SeUrl -Url $url_login -Driver $Driver

    # 填入帳號密碼,按登入
    $driver.FindElementByXPath("//input[@name='cn']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pw']").SendKeys($password)
    Send-SeKeys -Keys 'Return' -Driver $Driver
    
    # 日期轉換為民國年+月+日, 例如1130926
    # 取得今天的日期
    $Today = Get-Date
    # 計算民國年
    $TaiwanYear = $Today.Year - 1911
    # 組合民國年、月、日
    $TaiwanDate = "{0:D3}{1:D2}{2:D2}" -f $TaiwanYear, $Today.Month, $Today.Day

    # 登入完, 用powershell發出requests 模擬post 查詢報表
    $result = Invoke-WebRequest -uri $url_query -Method POST -Body "g_yyymmdd_s=$TaiwanDate&from=$branch" -SessionVariable session

    # 網頁存檔
     $result.Content | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_$($date).html"

     Enter-SeUrl -Url "$($save_path)\$($check_item)_$($branch)_$($date).html" -Driver $driver

    $width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    $height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # 調整視窗大小, 用以全螢幕截圖
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(($width + 120), ($height+200))
    # 截圖存檔
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # 關閉瀏覽器 
    Stop-SeDriver -Driver $driver

    # 回傳結果
    $result = @{"check_item" = $check_item;
    "branch"             = $branch;
    "png_filepath"       = "$($save_path)\$($check_item)_$($branch)_$($date).png";
    "html_filepath"      = "$($save_path)\$($check_item)_$($branch)_$($date).html"
    }
    return $result
}

#check-oe -check_item 'cpoe' -branch 'vhwc' -url 'http://172.20.200.71/cpoe/m2/batch' -account 'CC4F' -password 'acervghtc' -save_path 'd:\mis'

# 取得日期, 以便命名檔案
$date = (get-date).ToString('yyyyMMddhhmm')


foreach ($key in $check_oe.keys) {
    check-oe -check_item $key -branch $check_cpoe[$key]['branch'] -url $check_cpoe[$key]['url'] -account $check_cpoe[$key]['account'] -password $check_cpoe[$key]['password'] -save_path 'd:\mis' 
}

foreach ($key in $check_showjob.keys) {
    check-showjob -check_item $key -branch $check_showjob[$key]['branch'] -url $check_showjob[$key]['url'] -save_path 'd:\mis'
}