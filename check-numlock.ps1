#啟用NumLock
<#
解決Windows 10 開機預設都沒有啟動數字鍵(NumLock)的兩個方法。
一、	修改註冊檔
1、	Win+R輸入regedit開啟登陸編輯器。
2、	找到 \HKEY_USERS\.DEFAULT\Control Panel\Keyboard。
3、	將 InitialKeyboardIndicators 修改為80000002。
4、	重開機即可。
二、重新開機後再登入畫面先開啟NumLock，不登入直接重開機一次即可。
#>

param($runadmin)
function check-numlock {
    #先檢查是否己使用機碼修改方式啟用

    #建立連到HKEY_USERS的路徑
    New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -ErrorAction SilentlyContinue

    $is_numlock = Get-ItemProperty -Path 'HKU:\.DEFAULT\Control Panel\Keyboard' -Name "InitialKeyboardIndicators" 

    if ($is_numlock.InitialKeyboardIndicators -ne 80000002) {

        if ($check_admin) {
            Write-Output "NunLock末啟用, 現在正在啟用..."
            Set-ItemProperty -Path 'HKU:\.DEFAULT\Control Panel\Keyboard' -Name "InitialKeyboardIndicators" -Value 80000002
        } else {
            Write-Warning "沒有系統管理員權限,且NumLock末啟用,請以系統管理員身分重新嘗試."
        }
    } else {
        Write-Output "NunLock己啟用."
    }

    Remove-PSDrive -Name HKU
}


 
 
#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    check-numlock
    pause
}
