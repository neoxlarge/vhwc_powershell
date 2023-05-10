#設定adboe pdf不自動更新

param($runadmin)


<#
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown]
"bUpdateToSingleApp"=dword:00000000
#>

function check-pdf {

    $reg_path = "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown"
    $is_update = Get-ItemProperty -Path $reg_path -Name "bUpdateToSingleApp" -ErrorAction SilentlyContinue

    if ($is_update.bUpdateToSingleApp -eq 0) {
        Write-Output "Adobe PDF Reader 己設定不自動更新."
    }
    else {
        if ($check_admin) {
            Write-Output "正在設定Adobe PDF Reader 為不自動新更中."
            Set-ItemProperty -Path $reg_path -Name "bUpdateToSingleApp" -Value 00000000 -type DWord -Force
        }
        else {
            Write-Warning "沒有系統管理員權限,且Adobe PDF Reader未設定不自動新,請以系統管理員身分重新嘗試."
        }
    }

}


 
#檔案獨立執行時會執行函式, 如果是被?入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    check-pdf
    pause
}