<#
selenium �ϥΤ�k
https://github.com/adamdriscoll/selenium-powershell
https://www.zenrows.com/blog/selenium-powershell#interaction-automation

chromedrive.exe �U����m
https://googlechromelabs.github.io/chrome-for-testing/#stable


#>


# �x�s�I�ϩM�����ɪ����|
$result_path = "\\172.20.5.185\mis\dutycheck_result"

# chromedrive ���|, ��powershell�w�p�|�컷�ݮୱ�D���W����, �ҥH�n�T�{���ݥD���W�����|.
# �w�p�O��b d:\mis\vhwc_powershell\chromedriver.exe
$chromedriver_path = "d:\mis\vhwc_powershell\"

# selenium module path
# ���ݮୱ�D�����w��selenium powershell �Ҳ�, �ζפJ���覡���J�Ҳ�.
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

    # �}���s����, headless �Ҧ�
    $driver = Start-SeChrome -WebDriverDirectory $chromedriver_path -headless
    # �}�Һ��}
    Enter-SeUrl -Url $url -Driver $Driver

    # ��J�b���K�X,���n�J
    $driver.FindElementByXPath("//input[@name='login']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pass']").SendKeys($password)
    $driver.FindElementByXPath("//input[@name='m2Login_submit']").Click()
    Start-Sleep -Seconds 3

    # �վ�����j�p, �ΥH���ù��I��
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1080)

    # �qcapture_area �P�_,�����n���ʨ쪺��m, �Hsendkey home, end ��@
    $body = $driver.FindElementByTagName("body")
    $body.SendKeys([OpenQA.Selenium.Keys]::Control + [OpenQA.Selenium.Keys]::$capture_area) 
   
    Start-Sleep -Seconds 1
    
    # �x�s�����M�I��
    $driver.PageSource | Out-File -FilePath "$($result_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile("$($result_path)\$($check_item)_$($branch)_$($date).png", "png")

    # �����s���� 
    Stop-SeDriver -Driver $Driver

    # �^�ǵ��G
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

    # �}���s����, headless �Ҧ�
    $driver = Start-SeChrome -WebDriverDirectory $chromedriver_path -headless
    # �}�Һ��}
    Enter-SeUrl -Url $url -Driver $Driver

    $driver.FindElementByXPath("//input[@id='btnExec']").Click()
    start-sleep -second 2

    # �վ�����j�p, �ΥH���ù��I��
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1080)
    
    # �qcapture_area �P�_,�����n���ʨ쪺��m, �Hsendkey home, end ��@
    $body = $driver.FindElementByTagName("body")
    $body.SendKeys([OpenQA.Selenium.Keys]::Control + [OpenQA.Selenium.Keys]::$capture_area) 
    start-sleep -second 1
    
    # �x�s�����M�I��
    $driver.PageSource | Out-File -FilePath "$($result_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile( "$($result_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # �����s���� 
    Stop-SeDriver -Driver $driver

    # �^�ǵ��G
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

    # ����t�έn���n�J,�~��d��
    # �}���s����, headless �Ҧ�
    $driver = Start-SeChrome -WebDriverDirectory $chromedriver_path -headless
    # �}�Һ��}
    Enter-SeUrl -Url $url_login -Driver $Driver

    # ��J�b���K�X,���n�J
    $driver.FindElementByXPath("//input[@name='cn']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pw']").SendKeys($password)
    $driver.FindElementByXPath("//input[@type='submit']").Click()
    
    
    # ����ഫ������~/��/��, �Ҧp113/09/26
    # ���o���Ѫ����
    $Today = Get-Date
    # �p�����~
    $TaiwanYear = $Today.Year - 1911
    # �զX����~�B��B��
    $TaiwanDate = "{0:D3}/{1:D2}/{2:D2}" -f $TaiwanYear, $Today.Month, $Today.Day
   
    # �n�J��, ��powershell�o�Xrequests ����post �d�߳���
    $body = @{
        'g_yyymmdd_s' = $TaiwanDate
        'from'        = $branch
    }
    
    # ����d�߷|��@�Ǯɶ��]
    $result = Invoke-WebRequest -uri $url_query -Method POST -Body $body
    Start-Sleep -Seconds 2

    # �����s�ɫ�AŪ��
    $result.Content | Out-File -FilePath "$($result_path)\$($check_item)_$($branch)_$($date).html"
    Enter-SeUrl -Url "$($result_path)\$($check_item)_$($branch)_$($date).html" -Driver $driver

    # ���o�����j�p, �Ÿq�������, �ȺI����.
    $width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    $height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # �վ�����j�p, �ΥH���ù��I��
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(($width), ($height))
    # �I�Ϧs��
    $driver.GetScreenshot( ).SaveAsFile( "$($result_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # �����s���� 
    Stop-SeDriver -Driver $driver

    # �^�ǵ��G
    $result = @{"check_item" = $check_item;
        "branch"             = $branch;
        "png_filepath"       = "$($result_path)\$($check_item)_$($branch)_$($date).png";
        "html_filepath"      = "$($result_path)\$($check_item)_$($branch)_$($date).html"
    }
    return $result
}


# ���o���, �H�K�R�W�ɮ�
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