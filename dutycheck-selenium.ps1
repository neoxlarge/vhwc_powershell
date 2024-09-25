<#
selenium 使用方法
https://github.com/adamdriscoll/selenium-powershell
https://www.zenrows.com/blog/selenium-powershell#interaction-automation




#>

Import-Module selenium

$check_cpoe = @{
    'vhwc_cpoe' = @{
        'branch' = "vhwc"
        'url' = 'http://172.20.200.71/cpoe/m2/batch'
        'account' = 'vhwcmis'
        'password' = 'Mis20190610'
    }
    'vhcy_cpoe' = @{
        'branch' = "vhcy"
        'url' = 'http://172.19.200.71/cpoe/m2/batch'
        'account' = 'vhwcmis'
        'password' = 'Mis20190610'
    }
}


function check-oe( $branch, $url, $account, $password ) {


    $driver = Start-SeChrome -WebDriverDirectory ".\" -headless

    Enter-SeUrl -Url $url -Driver $Driver

    $driver.GetScreenshot( ).SaveAsFile( "screenshotxxx.png", "png" )

    Stop-SeDriver -Driver $Driver

}

$driver = Start-SeChrome -WebDriverDirectory ".\" -headless 



Enter-SeUrl 'https://www.google.com.tw/' -Driver $Driver

$driver
#$Html = $Driver.PageSource
#$Html

$driver.GetScreenshot( ).SaveAsFile( "screenshotxxx.png", "png" )
# close the browser and release its resources
Stop-SeDriver -Driver $Driver