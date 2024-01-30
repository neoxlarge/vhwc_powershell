#清除DNS快取
ipconfig /flushdns
#ipconfig /renew
Clear-DnsClientCache

#停止IE執行
Get-Process -Name iexplore -ErrorAction SilentlyContinue | Stop-Process -Force

#刪掉offline data
Get-ChildItem C:\2100\SSO\OFFLINEDATA | Remove-Item -Recurse -Force

   
<#
    清除Internet Explorer暫存檔和Cookies
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
修改IE內的DNS快取, 預設IE內的DNS快取會有30秒
https://support.microsoft.com/en-us/topic/how-internet-explorer-uses-the-cache-for-dns-host-entries-33d93ec9-e4fa-1557-4e9c-83517fed474f
#>

$ie_dnscachetime = @{
path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
name = "DnsCacheTimeout"
value = 3
PropertyType = "DWord"
Force = $true
}

$ie_serverinfotimeout =  @{
    path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings"
    name = "ServerInfoTimeOut"
    value = 300
    PropertyType = "DWord"
    Force = $true
}

New-ItemProperty @ie_dnscachetime
New-ItemProperty @ie_serverinfotimeout

#執行公文環境檔
$pathfile = "\\172.20.5.187\mis\08-2100公文系統\01公文環境檔.exe"
if (Test-Path $pathfile) {
    Copy-Item -Path $pathfile C:\2100 -Force
    & C:\2100\01公文環境檔.exe
}