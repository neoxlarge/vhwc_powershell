<#
selenium 使用方法
https://github.com/adamdriscoll/selenium-powershell
https://www.zenrows.com/blog/selenium-powershell#interaction-automation




#>

Import-Module selenium

$check_oe = @{
    'vhwc_cpoe' = @{
        'check_item' = 'cpoe'
        'branch' = "vhwc"
        'url' = 'http://172.20.200.71/cpoe/m2/batch'
        'account' = 'CC4F'
        'password' = 'acervghtc'
        'capture_area' = 'botten'
    };
    'vhcy_cpoe' = @{
        'check_item' = 'cpoe'
        'branch' = "vhcy"
        'url' = 'http://172.19.200.71/cpoe/m2/batch'
        'account' = 'CC4F'
        'password' = 'acervghtc'
        'capture_area' = 'botten'
    };

    'vhwc_eroe' = @{
        'check_item' = 'eroe'
        'branch' = "vhwc"
        'url' = 'http://172.20.200.71/eroe/m2/batch'
        'account' = 'CC4F'
        'password' = 'acervghtc'
        'capture_area' = 'top'
    };
    'vhcy_eroe' = @{
        'check_item' = 'eroe'
        'branch' = "vhcy"
        'url' = 'http://172.19.200.71/eroe/m2/batch'
        'account' = 'CC4F'
        'password' = 'acervghtc'
        'capture_area' = 'top'
    };
    }  
}


function check-oe( $check_item, $branch, $url, $account, $password, $save_path ) {
    # 取得日期, 以便命名檔案
    $date = (get-date).ToString('yyyyMMddhhmm')

    # 開啟瀏覽器, headless 模式
    $driver = Start-SeChrome -WebDriverDirectory ".\" -headless
    # 開啟網址
    Enter-SeUrl -Url $url -Driver $Driver

    # 填入帳號密碼,按登入
    $driver.FindElementByXPath("//input[@name='login']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pass']").SendKeys($password)
    $driver.FindElementByXPath("//input[@name='m2Login_submit']").Click()

    # 取得登入後的頁面的大小
    $width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    $height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # 調整視窗大小, 用以全螢幕截圖
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(($width + 120), ($height+200))
    
    # 儲存頁面和截圖
    $driver.PageSource | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # 關閉瀏覽器 
    Stop-SeDriver -Driver $Driver

    # 回傳結果
    $result = @{"check_item" = $check_item; 
                "branch" = $branch; 
                "date" = $date; 
                "png_filepath" = "$($save_path)\$($check_item)_$($branch)_$($date).png";
                "html_filepath" = "$($save_path)\$($check_item)_$($branch)_$($date).html"
            }

    return $result        

}

$check_showjob = @{
    'vhwc_showjob' = @{
        'check_item' ='showjob'
        'branch' = "vhwc"
        'url' = 'http://172.20.200.41/NOPD/showjoblog.aspx'
        'capture_area' = 'top'
    }
    
    'vhcy_showjob' = @{
        'check_item' ='showjob'
        'branch' = "vhcy"
        'url' = 'http://172.19.200.41/NOPD/showjoblog.aspx'
        'capture_area' = 'top'
    }
        
}

function check-showjob ($check_item,$branch, $url,$save_path) {

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
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1024)
    
    #  從capture_area 判斷,網頁要捲動到的位置
    $scroll_position = 0
    switch ($check_oe[$key]['capture_area']) {
        'top' { $scroll_position = 0 }
        'botten' { $scroll_position = $driver.ExecuteScript("return document.body.scrollHeight - window.innerHeight") }
     }

    # 捲動到特定位置
    $driver.ExecuteScript("window.scrollTo(0, $scroll_position)")

    start-sleep -second 5}
    
    # 儲存頁面和截圖
    $driver.PageSource | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_showjob.html"
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_showjob.png", "png" )

    # 關閉瀏覽器 
    Stop-SeDriver -Driver $driver

    # 回傳結果
    $result = @{"check_item" = $check_item; 
                "branch" = $branch; 
                "png_filepath" = "$($save_path)\$($check_item)_$($branch)_showjob.png";
                "html_filepath" = "$($save_path)\$($check_item)_$($branch)_showjob.html"
            }
    return $result        
}





#check-oe -check_item 'cpoe' -branch 'vhwc' -url 'http://172.20.200.71/cpoe/m2/batch' -account 'CC4F' -password 'acervghtc' -save_path 'd:\mis'

foreach ($key in $check_oe.keys) {
    check-oe -check_item $key -branch $check_cpoe[$key]['branch'] -url $check_cpoe[$key]['url'] -account $check_cpoe[$key]['account'] -password $check_cpoe[$key]['password'] -save_path 'd:\mis' 
}