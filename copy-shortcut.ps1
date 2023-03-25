#復制捷徑到公用桌面c:\user\public\desktop
#公用桌面內的捷徑,一般使用者帳號無法刪除, 可以避免誤刪捷徑.
#執行時需要管理者帳號.

param($runadmin)

function copy-shortcut {
    $source_path = "\\172.20.1.14\update\0001-中榮系統環境設定\VHWC_桌面常用捷徑"
    $destination_path = $env:PUBLIC+"\desktop"   # " C:\Users\Public\Desktop\"
    
    if ($check_admin) {
        Write-Output "復制常用捷徑到公用桌面."
        Start-Process -FilePath robocopy.exe -ArgumentList ($source_path + " " + $destination_path + " /e" ) -wait 
    } else {
        Write-Warning "沒有系統管理員權限, 未復制常用捷徑到公用桌面."
    }
    
}

#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    copy-shortcut
    pause
}