<#check-smartwonder
SmartWonder是用瀏覽器看PACS的
1. 開敫smartwonder會要求寫入c:\program files\tedpc
2. 會下載安裝2個元件, 但憑證己過期, 要改IE選項.
3. 在非PACS允許的電腦上, 會開眼睛來看PACE, ultraquery 有2個選項要改.
4. IE和EDGE設定方不同,EDGE也是要用IE mode.

#>

param($runadmin)

#管理者權限vhwcmis的證書.
$Username = "vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)


function check_smartwonder {

    #1. check c:\program files\tedpc
    if (!(Test-Path -path "c:\program files\tedpc")) {
        New-Item -Path "c:\program files\tedpc" -ItemType Directory -Force
    }
    
    #grant control to the folder
    #留到grant-fullcontrolpermission再開權限.




}

#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Check-VncSetting
    Check-VncService

    pause
}
