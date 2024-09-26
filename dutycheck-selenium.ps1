<#
selenium �ϥΤ�k
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
    # ���o���, �H�K�R�W�ɮ�
    $date = (get-date).ToString('yyyyMMddhhmm')

    # �}���s����, headless �Ҧ�
    $driver = Start-SeChrome -WebDriverDirectory ".\" -headless
    # �}�Һ��}
    Enter-SeUrl -Url $url -Driver $Driver

    # ��J�b���K�X,���n�J
    $driver.FindElementByXPath("//input[@name='login']").SendKeys($account)
    $driver.FindElementByXPath("//input[@name='pass']").SendKeys($password)
    $driver.FindElementByXPath("//input[@name='m2Login_submit']").Click()

    # ���o�n�J�᪺�������j�p
    $width = $driver.ExecuteScript("return document.documentElement.scrollWidth")
    $height = $driver.ExecuteScript("return document.documentElement.scrollHeight")
    
    # �վ�����j�p, �ΥH���ù��I��
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(($width + 120), ($height+200))
    
    # �x�s�����M�I��
    $driver.PageSource | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_$($date).html"
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_$($date).png", "png" )

    # �����s���� 
    Stop-SeDriver -Driver $Driver

    # �^�ǵ��G
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
    $driver.Manage().Window.Size = New-Object System.Drawing.Size(1920, 1024)
    
    #  �qcapture_area �P�_,�����n���ʨ쪺��m
    $scroll_position = 0
    switch ($check_oe[$key]['capture_area']) {
        'top' { $scroll_position = 0 }
        'botten' { $scroll_position = $driver.ExecuteScript("return document.body.scrollHeight - window.innerHeight") }
     }

    # ���ʨ�S�w��m
    $driver.ExecuteScript("window.scrollTo(0, $scroll_position)")

    start-sleep -second 5}
    
    # �x�s�����M�I��
    $driver.PageSource | Out-File -FilePath "$($save_path)\$($check_item)_$($branch)_showjob.html"
    $driver.GetScreenshot( ).SaveAsFile( "$($save_path)\$($check_item)_$($branch)_showjob.png", "png" )

    # �����s���� 
    Stop-SeDriver -Driver $driver

    # �^�ǵ��G
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