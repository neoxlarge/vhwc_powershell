<#
selenium �ϥΤ�k
https://github.com/adamdriscoll/selenium-powershell
https://www.zenrows.com/blog/selenium-powershell#interaction-automation

chromedrive.exe �U����m
https://googlechromelabs.github.io/chrome-for-testing/#stable

# line token(�W���ˬd�s��): HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz
# line token(����1): CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI
# line token(����2): AVt3SxMcHhatY2fuG2j6HzKGdb5BOTmrfAlEiBolQOO
# �w�ɨC�ѱߤW11:20��, �M���W0�I20������.

#>


# �x�s�I�ϩM�����ɪ����|
$result_path = "d:\mis\dutycheck_result"

# chromedrive ���|, ��powershell�w�p�|�컷�ݮୱ�D���W����, �ҥH�n�T�{���ݥD���W�����|.
# �w�p�O��b d:\mis\vhwc_powershell\chromedriver.exe
$chromedriver_path = "d:\mis\vhwc_powershell"

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



function check-oe( $check_item, $branch, $url, $account, $password, $capture_area) {

    # �}���s����, headless �Ҧ�
    $driver = Start-SeChrome -WebDriverDirectory $chromedriver_path -headless 
    # �}�Һ��}
    Enter-SeUrl -Url $url -Driver $Driver
    write-debug "check oe: $check_item $branch"
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



function Convert-Html2Table ($htmlFilePath) {
    # �Nhtml�ɮפ���table�ഫ��hash table
    # �Ѽ�: $htmlFilePath: html�ɮת����|
    # �^��: hash table

    # Ū��HTML�ɮפ��e
    $html = Get-Content -Path $htmlFilePath -Raw -Encoding UTF8

    # �ϥΥ��h��F���ǰt��椺�e
    $tablePattern = "(?s)<table[^>]*>.*?</table>"
    $rowPattern = "(?s)<tr[^>]*>(.*?)</tr>"
    $cellPattern = "(?s)<t[hd][^>]*>(.*?)</t[hd]>"

    # ��ơG�M�zHTML���e
    function Clean-HtmlContent($content) {
        # �B�z�S���p�A�p <a> ����
        $content = [regex]::Replace($content, '<a[^>]*>(.*?)</a>', '$1')
        
        # ������LHTML����
        $content = $content -replace '<[^>]+>', ''
        
        # ����HTML����M�M�z�ť�
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
                # ���]�Ĥ@��O���Y
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
        [bool]$notificationDisabled = $true  # �]�m�q���O�_�T�Ϊ��ѼơA�w�]���T��
    )

    Add-Type -AssemblyName System.Net.Http

    $uri = "https://notify-api.line.me/api/notify"

    # �ǳưT�����e
    $body = @{
        message = $message
        notificationDisabled = $notificationDisabled  # �N notificationDisabled �ѼƲK�[��T�����e��
    }

    # �ǳ�multipart/form-data �榡�����e
    $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
    foreach ($key in $body.Keys) {
        $content = [System.Net.Http.StringContent]::new($body[$key])
        $multipartContent.Add($content, $key)
    }

    # �[�J�Ϥ�
    if ($imagePath -ne "") {
        $imageStream = [System.IO.File]::OpenRead($imagePath)
        $imageContent = [System.Net.Http.StreamContent]::new($imageStream)
        $imageContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("image/png")  # �i���u???��?��?��
        $multipartContent.Add($imageContent, "imageFile", (Split-Path $imagePath -Leaf))
    }

    # �ǳ�HTTP�ШD
    $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $uri)
    $request.Headers.Authorization = "Bearer $token"
    $request.Content = $multipartContent

    # �o�e�ШD
    $httpClient = [System.Net.Http.HttpClient]::new()
    $response = $httpClient.SendAsync($request).Result

    # �B�z�^��
    if ($response.IsSuccessStatusCode) {
        Write-Host "�T���o�e���\�C"
    }
    else {
        Write-Host "�L�k�o�e�T���CStatusCode: $($response.StatusCode)�A��]: $($response.ReasonPhrase)"
    }

    start-sleep -second 2
}



# ���o���, �H�K�R�W�ɮ�
$date = (get-date).ToString('yyyyMMddhhmm')


foreach ($key in $check_oe.keys) {
    $result = check-oe -check_item $check_oe[$key]['check_item'] -branch $check_oe[$key]['branch'] -url $check_oe[$key]['url'] -account $check_oe[$key]['account'] -password $check_oe[$key]['password'] -capture_area $check_oe[$key]['capture_area']
    
    # �o�eLINE�I��
    Send-LineNotify -message $result['check_item'] -imagePath $result['png_filepath']
    
    # �ˬd���~
    $result_table = (convert-html2table -htmlFilePath $result['html_filepath']).Table1
    
    # �⦳���~,���~,���Ѧr�ꪺ�O����X��
    $error_talbe = @()
    foreach ($table_item in $result_table) {
        
        if ( $table_item['���檬�A'] -match '���~|����|���~') {
            $table_item['���檬�A'] = $table_item['���檬�A'] -replace '<[^>]+>', ''  # �����Ҧ� HTML ����
            $table_item['���檬�A'] = $table_item['���檬�A'].Trim()  # ���������ť�
            $error_talbe += $table_item
        }
    
    }
    # �����~�~�o�eLINE�T��
    if ($error_talbe.Count -gt 0) { 
        foreach ($error_item in $error_talbe) {
            $error_message = "? Fail: $($result['check_item']) `n �u�@ID: $($error_item['�妸�u�@ID']) `n���檬�A: $($error_item['���檬�A']) `n�}�l�ɶ�: $($error_item['�}�l�ɶ�']) `n����: $($error_item['����'])"
            Send-LineNotify -message $error_message 
        }
    }

} 

foreach ($key in $check_showjob.keys) {
    $result = check-showjob -check_item $check_showjob[$key]['check_item'] -branch $check_showjob[$key]['branch'] -url $check_showjob[$key]['url']  -capture_area $check_showjob[$key]['capture_area']

    # �o�eLINE�I��
    Send-LineNotify -message $result['check_item'] -imagePath $result['png_filepath']

    # �ˬd���~
    $result_table = (convert-html2table -htmlFilePath $result['html_filepath']).Table1

    # �⦳���~,���~,���Ѧr�ꪺ�O����X��
    $error_talbe = @()
    foreach ($table_item in $result_table) {
        
        if ( $table_item['�����ɶ�'] -match '���~|����|���~') {
            #$table_item['�����ɶ�'] = $table_item['���檬�A'] -replace '<[^>]+>', ''  # �����Ҧ� HTML ����
            #$table_item['���檬�A'] = $table_item['���檬�A'].Trim()  # ���������ť�
            $error_talbe += $table_item
        }
    
    }

    if ($error_talbe.Count -gt 0) { 
        foreach ($error_item in $error_talbe) {
            $error_message = "? Fail: $($result['check_item']) `n �{���N�X: $($error_item['�{���N�X']) `n���A: $($error_item['�����ɶ�']) `n����ɶ�: $($error_item['����ɶ�']) `n����: $($error_item['���檬�p'])"
            Send-LineNotify -message $error_message 
        }
    }

}

foreach ($key in $check_cyp2001.keys) {
    $result = check-cyp2001 -check_item $check_cyp2001[$key]['check_item'] -branch $check_cyp2001[$key]['branch'] -account $check_cyp2001[$key]['account'] -password $check_cyp2001[$key]['password'] -url_login $check_cyp2001[$key]['url_login'] -url_query $check_cyp2001[$key]['url_query'] 

     # �o�eLINE�I��
     Send-LineNotify -message $result['check_item'] -imagePath $result['png_filepath']
}

