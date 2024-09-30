<#
selenium 使用方法
https://github.com/adamdriscoll/selenium-powershell
https://www.zenrows.com/blog/selenium-powershell#interaction-automation

chromedrive.exe 下載位置
https://googlechromelabs.github.io/chrome-for-testing/#stable

# line token(灣橋檢查群組): HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz
# line token(測試1): CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI
# line token(測試2): AVt3SxMcHhatY2fuG2j6HzKGdb5BOTmrfAlEiBolQOO
# 定時每天晚上11:20分, 和早上0點20分執行.

#>


# 儲存截圖和網頁檔的路徑
$result_path = "d:\mis\dutycheck_result"

# chromedrive 路徑, 此powershell預計會到遠端桌面主機上執行, 所以要確認遠端主機上的路徑.
# 預計是放在 d:\mis\vhwc_powershell\chromedriver.exe
$chromedriver_path = "d:\mis\vhwc_powershell"

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



function check-oe( $check_item, $branch, $url, $account, $password, $capture_area) {

    # 開啟瀏覽器, headless 模式
    $driver = Start-SeChrome -WebDriverDirectory $chromedriver_path -headless 
    # 開啟網址
    Enter-SeUrl -Url $url -Driver $Driver
    write-debug "check oe: $check_item $branch"
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



function Convert-Html2Table ($htmlFilePath) {
    # 將html檔案中的table轉換成hash table
    # 參數: $htmlFilePath: html檔案的路徑
    # 回傳: hash table

    # 讀取HTML檔案內容
    $html = Get-Content -Path $htmlFilePath -Raw -Encoding UTF8

    # 使用正則表達式匹配表格內容
    $tablePattern = "(?s)<table[^>]*>.*?</table>"
    $rowPattern = "(?s)<tr[^>]*>(.*?)</tr>"
    $cellPattern = "(?s)<t[hd][^>]*>(.*?)</t[hd]>"

    # 函數：清理HTML內容
    function Clean-HtmlContent($content) {
        # 處理特殊情況，如 <a> 標籤
        $content = [regex]::Replace($content, '<a[^>]*>(.*?)</a>', '$1')
        
        # 移除其他HTML標籤
        $content = $content -replace '<[^>]+>', ''
        
        # 替換HTML實體和清理空白
        $content = $content -replace '&nbsp;', ' ' `
                            -replace '&lt;', '<' `
                            -replace '&gt;', '>' `
                            -replace '&amp;', '&' `
                            -replace '^\s+|\s+$', '' `
                            -replace '\s+', ' '
        return $content
    }

    $result = @{}
    $tableMatches = [regex]::Matches($html, $tablePattern)

    for ($i = 0; $i -lt $tableMatches.Count; $i++) {
        $tableContent = $tableMatches[$i].Value
        $rows = [regex]::Matches($tableContent, $rowPattern)

        $headers = @()
        $tableData = @()

        for ($j = 0; $j -lt $rows.Count; $j++) {
            $rowContent = $rows[$j].Groups[1].Value
            $cells = [regex]::Matches($rowContent, $cellPattern)
            
            if ($j -eq 0) {
                # 假設第一行是表頭
                $headers = $cells | ForEach-Object { 
                    Clean-HtmlContent $_.Groups[1].Value
                }
            } else {
                $rowData = @{}
                for ($k = 0; $k -lt $headers.Count; $k++) {
                    if ($k -lt $cells.Count) {
                        $cellValue = Clean-HtmlContent $cells[$k].Groups[1].Value
                        $rowData[$headers[$k]] = $cellValue
                    } else {
                        $rowData[$headers[$k]] = $null
                    }
                }
                $tableData += $rowData
            }
        }

        $result["Table$($i+1)"] = $tableData
    }

    return $result
}


function Send-LineNotify {
    param (
        [string]$token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI",
        [string]$message,
        [string]$imagePath,
        [bool]$notificationDisabled = $true  # 設置通知是否禁用的參數，預設為禁用
    )

    Add-Type -AssemblyName System.Net.Http

    $uri = "https://notify-api.line.me/api/notify"

    # 準備訊息內容
    $body = @{
        message = $message
        notificationDisabled = $notificationDisabled  # 將 notificationDisabled 參數添加到訊息內容中
    }

    # 準備multipart/form-data 格式的內容
    $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
    foreach ($key in $body.Keys) {
        $content = [System.Net.Http.StringContent]::new($body[$key])
        $multipartContent.Add($content, $key)
    }

    # 加入圖片
    if ($imagePath -ne "") {
        $imageStream = [System.IO.File]::OpenRead($imagePath)
        $imageContent = [System.Net.Http.StreamContent]::new($imageStream)
        $imageContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("image/png")  # 可根据???片?型?整
        $multipartContent.Add($imageContent, "imageFile", (Split-Path $imagePath -Leaf))
    }

    # 準備HTTP請求
    $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $uri)
    $request.Headers.Authorization = "Bearer $token"
    $request.Content = $multipartContent

    # 發送請求
    $httpClient = [System.Net.Http.HttpClient]::new()
    $response = $httpClient.SendAsync($request).Result

    # 處理回應
    if ($response.IsSuccessStatusCode) {
        Write-Host "訊息發送成功。"
    }
    else {
        Write-Host "無法發送訊息。StatusCode: $($response.StatusCode)，原因: $($response.ReasonPhrase)"
    }

    start-sleep -second 2
}



# 取得日期, 以便命名檔案
$date = (get-date).ToString('yyyyMMddhhmm')


foreach ($key in $check_oe.keys) {
    $result = check-oe -check_item $check_oe[$key]['check_item'] -branch $check_oe[$key]['branch'] -url $check_oe[$key]['url'] -account $check_oe[$key]['account'] -password $check_oe[$key]['password'] -capture_area $check_oe[$key]['capture_area']
    
    # 發送LINE截圖
    Send-LineNotify -message $result['check_item'] -imagePath $result['png_filepath']
    
    # 檢查錯誤
    $result_table = (convert-html2table -htmlFilePath $result['html_filepath']).Table1
    
    # 把有錯誤,有誤,失敗字串的記錄選出來
    $error_talbe = @()
    foreach ($table_item in $result_table) {
        
        if ( $table_item['執行狀態'] -match '錯誤|失敗|有誤') {
            $table_item['執行狀態'] = $table_item['執行狀態'] -replace '<[^>]+>', ''  # 移除所有 HTML 標籤
            $table_item['執行狀態'] = $table_item['執行狀態'].Trim()  # 移除首尾空白
            $error_talbe += $table_item
        }
    
    }
    # 有錯誤才發送LINE訊息
    if ($error_talbe.Count -gt 0) { 
        foreach ($error_item in $error_talbe) {
            $error_message = "? Fail: $($result['check_item']) `n 工作ID: $($error_item['批次工作ID']) `n執行狀態: $($error_item['執行狀態']) `n開始時間: $($error_item['開始時間']) `n說明: $($error_item['說明'])"
            Send-LineNotify -message $error_message 
        }
    }

} 

foreach ($key in $check_showjob.keys) {
    $result = check-showjob -check_item $check_showjob[$key]['check_item'] -branch $check_showjob[$key]['branch'] -url $check_showjob[$key]['url']  -capture_area $check_showjob[$key]['capture_area']

    # 發送LINE截圖
    Send-LineNotify -message $result['check_item'] -imagePath $result['png_filepath']

    # 檢查錯誤
    $result_table = (convert-html2table -htmlFilePath $result['html_filepath']).Table1

    # 把有錯誤,有誤,失敗字串的記錄選出來
    $error_talbe = @()
    foreach ($table_item in $result_table) {
        
        if ( $table_item['結束時間'] -match '錯誤|失敗|有誤') {
            #$table_item['結束時間'] = $table_item['執行狀態'] -replace '<[^>]+>', ''  # 移除所有 HTML 標籤
            #$table_item['執行狀態'] = $table_item['執行狀態'].Trim()  # 移除首尾空白
            $error_talbe += $table_item
        }
    
    }

    if ($error_talbe.Count -gt 0) { 
        foreach ($error_item in $error_talbe) {
            $error_message = "? Fail: $($result['check_item']) `n 程式代碼: $($error_item['程式代碼']) `n狀態: $($error_item['結束時間']) `n執行時間: $($error_item['執行時間']) `n說明: $($error_item['執行狀況'])"
            Send-LineNotify -message $error_message 
        }
    }

}

foreach ($key in $check_cyp2001.keys) {
    $result = check-cyp2001 -check_item $check_cyp2001[$key]['check_item'] -branch $check_cyp2001[$key]['branch'] -account $check_cyp2001[$key]['account'] -password $check_cyp2001[$key]['password'] -url_login $check_cyp2001[$key]['url_login'] -url_query $check_cyp2001[$key]['url_query'] 

     # 發送LINE截圖
     Send-LineNotify -message $result['check_item'] -imagePath $result['png_filepath']
}

