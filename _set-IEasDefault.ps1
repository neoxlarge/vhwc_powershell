#把IE設成預設值.
#1. 關閉EDGE中的IE模式.
#2. 關閉IE中的IEtoEdga元件
#3. 把IE設成預設值, 這動作無法直接修改registry, 因為無法算出hash.
#   需要group policy的方式?成.

param($runadmin)

set-IEasDefault {

    Write-Output "在Edge中以IE開啟網站設為永不"
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Edge\IEToEdge" -Name "RedirectionMode" -Value 0 -Force

    Write-Output "關閉IE中的IEtoEdga元件"
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID" -Name "{1FD49718-1D00-4B19-AF5F-070AF6D5D54C}" -Value 0 -Force

    

}



#檔案獨立執行時會執行函式, 如果是被?入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    Set-HomePage
    pause
}
