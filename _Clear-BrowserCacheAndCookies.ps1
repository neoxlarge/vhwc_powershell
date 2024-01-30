#清理IE, EDGE, Chrome 的暫存和cookies.

param($runadmin)


function Clear-BrowserCacheAndCookies {
    # 清除Internet Explorer暫存檔和Cookies
    
    <#
    Delete Temporary Internet Files:
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8

    Delete Cookies:
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2

    Delete History:
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1

    Delete Form Data:
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16

    Delete Passwords:
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32

    Delete All:
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255

    Delete All + files and settings stored by Add-ons:
    RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351
    #>


    Start-Process -FilePath "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess 8" -Wait
    Start-Process -FilePath "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess 2" -Wait
   
  
    # 清除Google Chrome暫存檔和Cookies
    try {
      Remove-Item -Path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -ErrorAction  Stop
      Remove-Item -Path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Network\cookies" -Recurse -Force -ErrorAction Stop 
    }
    catch [System.Management.Automation.ItemNotFoundException] {
      Write-Warning "清除暫存檔可能失敗:"
      Write-Warning $Error[0].Exception.Message
    }

    
    # 清除Microsoft Edge暫存檔和Cookies
    try {
    Remove-Item -Path "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\Cache\*" -Recurse -Force -ErrorAction Stop
    Remove-Item -Path "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\Cookies\*" -Recurse -Force -ErrorAction Stop
    }
    catch [System.Management.Automation.ItemNotFoundException]{
      Write-Warning "清除暫存檔可能失敗:"
      Write-Warning $Error[0].Exception.Message
    }
  }

  
#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Clear-BrowserCacheAndCookies
    pause
}