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


#清除DNS快取
ipconfig /flushdns
#ipconfig /renew
#Clear-DnsClientCache

#停止IE執行
Get-Process -Name iexplore,chrome,msedge -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

  
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
# 定義要設置的註冊表項
$ie_settings = @(
    @{
        path         = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
        name         = "DnsCacheTimeout"
        value        = 3
        PropertyType = "DWord"
    },
    @{
        path         = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
        name         = "ServerInfoTimeOut"
        value        = 300
        PropertyType = "DWord"
    },
    @{
        path         = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
        name         = "DnsCacheTimeout"
        value        = 3
        PropertyType = "DWord"
    },
    @{
        path         = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
        name         = "ServerInfoTimeOut"
        value        = 300
        PropertyType = "DWord"
    }
)

foreach ($setting in $ie_settings) {
    # 檢查路徑是否存在
    if (!(Test-Path $setting.path)) {
        New-Item -Path $setting.path -Force | Out-Null
    }

    # 檢查值是否存在且正確
    $existingValue = Get-ItemProperty -Path $setting.path -Name $setting.name -ErrorAction SilentlyContinue

    if ($null -eq $existingValue -or $existingValue.$($setting.name) -ne $setting.value) {
        # 如果值不存在或不正確，則設置新值
        if ($setting.path.StartsWith("HKLM:") -and !$check_admin) {
            Write-Warning "需要管理員權限來修改 HKLM 註冊表項：$($setting.path)\$($setting.name)"
        } else {
            Set-ItemProperty @setting -Force
            Write-Host "已更新 $($setting.path)\$($setting.name) 為 $($setting.value)"
        }
    } else {
        Write-Host "$($setting.path)\$($setting.name) 已經正確設置為 $($setting.value)"
    }
}

# 切換DNS設定, 把第一個DNS改為1.12或1.11, 改完測一下公文(edsap.edoc.vghtc.gov.tw)的連線, 如果還是不正常會顯示仍有問題.
# 最後會重置DNS設定, 改成自動取得.
# 省略前面的代碼...

import-module_func -name "NetAdapter"
import-module_func -name "DnsClient"

$cimParams = @{
    ComputerName = $env:COMPUTERNAME
    Credential = $credential
}
$cimSession = New-CimSession @cimParams

$network = Get-NetAdapter -CimSession $cimSession | Where-Object ConnectorPresent

$dnsParams = @{
    InterfaceIndex = $network.ifIndex
    AddressFamily = 'IPv4'
    CimSession = $cimSession
}
$dnsSettings = Get-DnsClientServerAddress @dnsParams

$websites = @(
    'eip.vghtc.gov.tw',
    'lcs.vghb12.vhwc.gov.tw',
    'ehis.vghb12.vhwc.gov.tw',
    'edap.doc.vghtc.gov.tw'
)

function Test-Websites {
    $results = @{}
    foreach ($site in $websites) {
        $testResult = Test-Connection -ComputerName $site -Count 1 -Quiet
        $results[$site] = $testResult
        $status = if ($testResult) { "成功" } else { "失敗" }
        Write-Host "測試 $site : $status"
    }
    return $results
}

Write-Host "原始DNS設定下的測試結果："
$originalResults = Test-Websites

$newDns = if ($dnsSettings.ServerAddresses[0] -eq '172.19.1.11') {
    '172.19.1.12', '172.19.1.11'
} else {
    '172.19.1.11', '172.19.1.12'
}

Write-Host "正在更改DNS設定為 $($newDns -join ', ')..."
Set-DnsClientServerAddress -InterfaceIndex $network.ifIndex -ServerAddresses $newDns -CimSession $cimSession

Write-Host "新DNS設定下的測試結果："
$newResults = Test-Websites

$allSuccessful = $newResults.Values | Where-Object { $_ -eq $false } | Measure-Object | Select-Object -ExpandProperty Count
$allSuccessful = $allSuccessful -eq 0

if ($allSuccessful) {
    Write-Host "所有網站都能成功連接。正在重置DNS設定為自動取得..."
    Set-DnsClientServerAddress -InterfaceIndex $network.ifIndex -ResetServerAddresses -CimSession $cimSession
    Write-Host "DNS設定已重置為自動取得。"
} else {
    Write-Warning "部分網站無法正確連接!! 請連絡資訊室,謝謝."
    Write-Host "正在重置DNS設定為自動取得..."
    Set-DnsClientServerAddress -InterfaceIndex $network.ifIndex -ResetServerAddresses -CimSession $cimSession
    Write-Host "DNS設定已重置為自動取得。"
    Pause
}