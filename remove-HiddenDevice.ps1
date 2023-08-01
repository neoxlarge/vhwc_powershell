#從裝置管理員中移除?藏的設備
#
#只會移除ScmarCard ,SmartCardReader和SmartCardFilter這3個和讀卡機相關的?藏的設備.
#powershell V2無法執行, get-pnpdevice是V5的語法.
#
#win32_pnpentity中不會列出?藏的設備, 無法用win32_pnpentity來作.

param($runadmin)

#todo: 部分系統旳pnputil.exe沒有remove-device
function remove-HiddenDevice {

  $dev =  Get-PnpDevice | Where-Object -FilterScript {$_.Present -eq $false -and $_.Class -in ('SmartCard','SmartCardReader','SmartCardFilter')}

  foreach ($d in $dev) {
    pnputil /remove-device "$($d.instanceID)"
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

    remove-HiddenDevice
    pause
}