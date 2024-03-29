
param($runadmin)

function Set-ScreenSaver {
    
    param (
        [switch]$off
    )

    write-output "設定螢幕保護程式"
  
    # 設定特定的螢幕保護程式檔案路徑
    $screenSaverFilePath = "c:\screensaver.scr"
    $result = !(Test-Path -Path $screenSaverFilePath) -and $check_admin
    if ($result) {
    Copy-Item -Path "\\172.20.1.14\update\Vghtc_Update\ScreenSaver\ScreenSaver.scr" -Destination $screenSaverFilePath -Force 
    }
    
    if (!$off) { 
        # 設定螢幕保護程式
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaveActive -Value 1

        # 設定螢幕保護程式等待時間（以秒為單位）
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaveTimeOut -Value 900 -type string

        # 設定螢幕保護程式的密碼保護狀態（0表示禁用，1表示啟用）
        #Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaverIsSecure -Value 1

        # 設定特定的螢幕保護程式檔案路徑
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name SCRNSAVE.EXE -Value $screenSaverFilePath
        
    } else {

        # 設定螢幕保護程式
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaveActive -Value 0

        # 設定螢幕保護程式等待時間（以秒為單位）
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaveTimeOut -Value 0 -type string

        # 設定螢幕保護程式的密碼保護狀態（0表示禁用，1表示啟用）
        #Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaverIsSecure -Value 1

        # 設定特定的螢幕保護程式檔案路徑
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name SCRNSAVE.EXE -Value ""
    
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

    set-ScreenSaver
    pause
}
