Import-Module selenium

$driver = Start-SeChrome -WebDriverDirectory ".\" -headless

Enter-SeUrl 'https://www.google.com.tw/' -Driver $Driver

#$Html = $Driver.PageSource
#$Html

$driver.GetScreenshot( ).SaveAsFile( "screenshotxxx.png", "png" )
# close the browser and release its resources
Stop-SeDriver -Driver $Driver