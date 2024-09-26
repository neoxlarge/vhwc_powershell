<#
selenium �ϥΤ�k
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

    # �}���s����, headless �Ҧ�
    $driver = Start-SeChrome -WebDriverDirectory ".\" -headless
    # �}�Һ��}
    Enter-SeUrl -Url $url -Driver $Driver

    # ��J�b���K�X,���n�J
    $driver.FindElementByXPath("//input[@name='login']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pass']").SendKeys($password)
    $driver.FindElementByXPath("//input[@name='m2Login_submit']").Click()
    Start-Sleep -Seconds 3

    # ���o�n�J�᪺�������j�p
    #$width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    #$height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # �վ�����j�p, �ΥH���ù��I��
    #$driver.Manage().Window.Size = New-Object System.Drawing.Size(($width + 120), ($height+200))
    driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1080)

    # �qcapture_area �P�_,�����n���ʨ쪺��m, �Hsendkey home, end ��@
    Send-SeKeys -Keys $check_oe[$key]['capture_area'] -Driver $Driver
    Start-Sleep -Seconds 2
    
    # �x�s�����M�I��
    $driver.PageSource | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # �����s���� 
    Stop-SeDriver -Driver $Driver

    # �^�ǵ��G
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

    # �}���s����, headless �Ҧ�
    $driver = Start-SeChrome -WebDriverDirectory ".\" -headless
    # �}�Һ��}
    Enter-SeUrl -Url $url -Driver $Driver

    $driver.FindElementByXPath("//input[@id='btnExec']").Click()

    start-sleep -second 5

    # ���o�n�J�᪺�������j�p
    #$width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    #$height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # �վ�����j�p, �ΥH���ù��I��
    #$driver.Manage().Window.Size = New-Object System.Drawing.Size(($width + 120), ($height+200))
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1080)
    
    # �qcapture_area �P�_,�����n���ʨ쪺��m, �Hsendkey home, end ��@
    Send-SeKeys -Keys $check_showjob[$key]['capture_area'] -Driver $Driver
    start-sleep -second 2
    
    # �x�s�����M�I��
    $driver.PageSource | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # �����s���� 
    Stop-SeDriver -Driver $driver

    # �^�ǵ��G
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

    # ����t�έn���n�J,�~��d��
    # �}���s����, headless �Ҧ�
    $driver = Start-SeChrome -WebDriverDirectory ".\" -headless
    # �}�Һ��}
    Enter-SeUrl -Url $url_login -Driver $Driver

    # ��J�b���K�X,���n�J
    $driver.FindElementByXPath("//input[@name='cn']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pw']").SendKeys($password)
    Send-SeKeys -Keys 'Return' -Driver $Driver
    
    # ����ഫ������~+��+��, �Ҧp1130926
    # ���o���Ѫ����
    $Today = Get-Date
    # �p�����~
    $TaiwanYear = $Today.Year - 1911
    # �զX����~�B��B��
    $TaiwanDate = "{0:D3}{1:D2}{2:D2}" -f $TaiwanYear, $Today.Month, $Today.Day

    # �n�J��, ��powershell�o�Xrequests ����post �d�߳���
    $result = Invoke-WebRequest -uri $url_query -Method POST -Body "g_yyymmdd_s=$TaiwanDate&from=$branch" -SessionVariable session

    # �����s��
     $result.Content | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_$($date).html"

     Enter-SeUrl -Url "$($save_path)\$($check_item)_$($branch)_$($date).html" -Driver $driver

    $width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    $height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # �վ�����j�p, �ΥH���ù��I��
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(($width + 120), ($height+200))
    # �I�Ϧs��
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # �����s���� 
    Stop-SeDriver -Driver $driver

    # �^�ǵ��G
    $result = @{"check_item" = $check_item;
    "branch"             = $branch;
    "png_filepath"       = "$($save_path)\$($check_item)_$($branch)_$($date).png";
    "html_filepath"      = "$($save_path)\$($check_item)_$($branch)_$($date).html"
    }
    return $result
}

#check-oe -check_item 'cpoe' -branch 'vhwc' -url 'http://172.20.200.71/cpoe/m2/batch' -account 'CC4F' -password 'acervghtc' -save_path 'd:\mis'

# ���o���, �H�K�R�W�ɮ�
$date = (get-date).ToString('yyyyMMddhhmm')


foreach ($key in $check_oe.keys) {
    check-oe -check_item $key -branch $check_cpoe[$key]['branch'] -url $check_cpoe[$key]['url'] -account $check_cpoe[$key]['account'] -password $check_cpoe[$key]['password'] -save_path 'd:\mis' 
}

foreach ($key in $check_showjob.keys) {
    check-showjob -check_item $key -branch $check_showjob[$key]['branch'] -url $check_showjob[$key]['url'] -save_path 'd:\mis'
}