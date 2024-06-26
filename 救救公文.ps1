#檢查是否管理員
$check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")


function import-module_func ($name) {

    $result = get-module -ListAvailable $name

    if ($result -ne $null) {
        Import-Module -Name $name -ErrorAction Stop
    }
    else {
        $credential = get-admin_cred
        $rsession = New-PSSession -ComputerName wcdc2.vhcy.gov.tw -Credential $credential
        Import-Module $name -PSSession $rsession -ErrorAction Stop
    }
}

Import-Module -name "$(Split-Path $PSCommandPath)\vhwcmis_module.psm1"


#清除DNS快取
ipconfig /flushdns
#ipconfig /renew
#Clear-DnsClientCache

#停止IE執行
Get-Process -Name iexplore -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

#刪掉公文系統的offline data

$Username = "vhcy\vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

if (test-path -Path "C:\2100\sso\OFFLINEDATA") {
    New-PSDrive -Name "del" -Root "C:\2100\SSO\OFFLINEDATA" -Credential $credential -PSProvider FileSystem
    Get-ChildItem "del:\" | Remove-Item -Recurse -Force 
    Remove-PSDrive -Name "del" 
}
   
<# 清除Internet Explorer暫存檔和Cookies
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


<#
修改IE內的DNS快取, 預設IE內的DNS快取會有30分鐘
https://support.microsoft.com/en-us/topic/how-internet-explorer-uses-the-cache-for-dns-host-entries-33d93ec9-e4fa-1557-4e9c-83517fed474f
#>

$ie_dnscachetime_hklm = @{
    path         = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
    name         = "DnsCacheTimeout"
    value        = 3
    PropertyType = "DWord"
    Force        = $true
}

$ie_serverinfotimeout_hklm = @{
    path         = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
    name         = "ServerInfoTimeOut"
    value        = 300
    PropertyType = "DWord"
    Force        = $true
}

$ie_dnscachetime_hkcu = @{
    path         = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
    name         = "DnsCacheTimeout"
    value        = 3
    PropertyType = "DWord"
    Force        = $true
}

$ie_serverinfotimeout_hkcu = @{
    path         = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
    name         = "ServerInfoTimeOut"
    value        = 300
    PropertyType = "DWord"
    Force        = $true
}

if ($check_admin) {
    New-ItemProperty @ie_dnscachetime_hklm 
    New-ItemProperty @ie_serverinfotimeout_hklm
}

New-ItemProperty @ie_dnscachetime_hkcu 
New-ItemProperty @ie_serverinfotimeout_hkcu

#執行公文環境檔
$pathfile = "\\172.20.5.187\mis\08-2100公文系統\01公文環境檔.exe"
if (Test-Path $pathfile) {
    Copy-Item -Path $pathfile C:\2100 -Force
    & C:\2100\01公文環境檔.exe
}


# 切換DNS設定, 把第一個DNS改為1.12或1.11, 改完測一下公文(edsap.edoc.vghtc.gov.tw)的連線, 如果還是不正常會顯示仍有問題.
# 最後會重置DNS設定, 改成自動取得.

import-module_func -name "NetAdapter"
import-module_func -name "DnsClient"

$cim_session = New-CimSession -ComputerName $env:COMPUTERNAME -Credential $credential

$network = Get-NetAdapter -CimSession $cim_session | Where-Object -FilterScript {$_.ConnectorPresent -eq $true}

$dns_setting = Get-DnsClientServerAddress -InterfaceIndex $network.ifIndex -AddressFamily IPv4 -CimSession $cim_session

if ($dns_setting.Address[0] -eq "172.19.1.11") {
    Set-DnsClientServerAddress -InterfaceIndex $network.ifIndex -ServerAddresses ("172.19.1.12","172.19.1.11") -CimSession $cim_session
} else {
    Set-DnsClientServerAddress -InterfaceIndex $network.ifIndex -ServerAddresses ("172.19.1.11","172.19.1.12") -CimSession $cim_session
}

if (Test-Connection -ComputerName "edsap.edoc.vghtc.gov.tw" -Quiet -Count 1) {
    Set-DnsClientServerAddress -InterfaceIndex $network.ifIndex -ResetServerAddresses -CimSession $cim_session
} else {
    Write-Warning "仍無法正確ping到公文系統(edsap.edoc.vghtc.gov.tw)!! 請連絡資訊室,謝謝."
    Pause
}

