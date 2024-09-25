<#
selenium 使用方法
https://github.com/adamdriscoll/selenium-powershell
https://www.zenrows.com/blog/selenium-powershell#interaction-automation




#>

Import-Module selenium

$driver = Start-SeChrome -WebDriverDirectory ".\" -headless 

$driver | gm

Enter-SeUrl 'https://www.google.com.tw/' -Driver $Driver

$driver
#$Html = $Driver.PageSource
#$Html

$driver.GetScreenshot( ).SaveAsFile( "screenshotxxx.png", "png" )
# close the browser and release its resources
Stop-SeDriver -Driver $Driver