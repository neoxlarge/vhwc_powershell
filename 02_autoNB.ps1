# 自動安裝和設定
param($run_admin)
$run_main = $true


# 以管理員權限執行, Elevate Powershell to Admin
$check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (!$check_admin -and !$run_admin) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -run_admin 1" -Verb RunAs; exit
}


#啟用記錄
#存log 檔的資料?
if ((Test-Path -Path "d:\")) {
    $log_folder = "d:\mis"
} else {$log_folder = "c:\mis"}
Start-Transcript -Path "$log_folder\$(Get-Date -Format 'yyyy-MM-dd-hh-mm').log" -Append

#記錄powershell是用什麼權限執行的
if ($check_admin) {
    Write-Output "Powershell run as Admin mode."
}
else {
    Write-Output "Powershell run as User mode."
}

#檢查powershell版本, 應該高於5.1, Win7預載為2.0, Win10預載為5.1, Win7需要安裝5.1
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "Powershell 版本應該要5.1以上, 中止執行." -ErrorAction Stop
}

#==安裝================================================================== 

#檢查及啟用SMBv1/CIFS功能. 此功能本應該重開機,但先取消重開機,等全部軟體裝完.
import-module ((split-path $PSCommandPath) + "\Enable-Win10OPTFeature.ps1")
Enable-SMBv1
Enable-NetFx3

# 復制VGHTC等醫院專用程式.
Import-Module ((Split-Path $PSCommandPath) + "\copy-vghtc.ps1")
copy-vghtc

# install Oracle 9i Client and BDE
Import-Module ((Split-Path $PSCommandPath) + "\install-or9iClient.ps1")
Import-Module ((Split-Path $PSCommandPath) + "\remove-or9iandBDE.ps1")

install-or9iclient
install-BDE

# install 7z
Import-Module ((Split-Path $PSCommandPath) + "\install-7z.ps1")
install-7Z

# install EZUSB
Import-Module ((Split-Path $PSCommandPath) + "\install-EZUSB.ps1")
install-EZUSB

## 安裝CMS_CGServiSignAdapter
### 依文件要求,安裝前應關閉防毒軟體, 所以比防毒先安裝
Import-Module ((Split-Path $PSCommandPath) + "\install-CMS.ps1")
install-CMS


## 安裝 HCAServiSign
Import-Module ((Split-Path $PSCommandPath) + "\install-HCA.ps1")
install-HCA

# 安裝Hicos
Import-Module ((Split-Path $PSCommandPath) + "\install-HiCOS.ps1")
install-Hicos

# 安裝IDC
Import-Module ((Split-Path $PSCommandPath) + "\install-IDC.ps1")
install-IDC

#安裝VNC
Import-Module ((Split-Path $PSCommandPath) + "\install-VNC.ps1")
install-vnc

#install java
Import-Module ((Split-Path $PSCommandPath) + "\install-Java.ps1")
install-java
set-Java_env

#安裝2100
Import-Module ((Split-Path $PSCommandPath) + "\install-2100.ps1")
install-2100
set-2100_env

#install-IE11
Import-Module ((Split-Path $PSCommandPath) + "\install-IE11.ps1")
install-IE11

#install chrome
Import-Module ((Split-Path $PSCommandPath) + "\install-chrome.ps1")
install-chrome

#install Edge and Webview
Import-Module ((Split-Path $PSCommandPath) + "\install-Edge.ps1")
Install-Edge
install-EdgeWebview

#install smartiris
Import-Module ((Split-Path $PSCommandPath) + "\install-smartiris.ps1")
#install-smartiris

#install pdf
Import-Module ((Split-Path $PSCommandPath) + "\install-pdf.ps1")
install-pdf
check-pdf


#install libreoffice
Import-Module ((Split-Path $PSCommandPath) + "\install-libreoffice.ps1")
install-libreoffice

# 安裝雲端安控元件健保卡讀卡機控制(PCSC)
Import-Module ((Split-Path $PSCommandPath) + "\install-PCSC.ps1")
install-PCSC

#安裝虛擬健保卡
Import-Module ((Split-Path $PSCommandPath) + "\install-virtualnhc.ps1")
install-virtualnhc

#安裝庫賈氏病勾稽查詢系統
Import-Module ((Split-Path $PSCommandPath) + "\install-cdcalert.ps1")
install-cdcalert


# 安裝sslvpn
Import-Module ((Split-Path $PSCommandPath) + "\install-sslvpn.ps1")
install-SMAConnectAgent
install-NX

#安裝anydesk
 Import-Module ((Split-Path $PSCommandPath) + "\install-AnyDesk.ps1")
 install-AnyDesk

# 安裝Winnexus
Import-Module ((Split-Path $PSCommandPath) + "\install-Winnexus.ps1")

#install-winNexus    

# 安裝防毒 Trend Micro Apex One Security Agent
Import-Module ((Split-Path $PSCommandPath) + "\install-AntiVir.ps1")

install-antivir

# 移除不必要的win10 程式
Import-Module ((Split-Path $PSCommandPath) + "\remove-apps.ps1")
remove-apps


####系統設定存######################################################################################

#啟用powershell遠端管理
import-module ((split-path $PSCommandPath) + "\Check-EnablePSRemoting.ps1")
#check-enablepsremoting

##修改資料的權限
import-module ((split-path $PSCommandPath) + "\grant-FullControlPermission.ps1")
Grant-FullControlPermission

#變更電源計畫
import-module ((split-path $PSCommandPath) + "\disable-sleep.ps1")
#disable-sleep

#關閉UAC
Import-Module ((Split-Path $PSCommandPath) + "\Disable-UAC.ps1")
Disable-UAC

#復制捷徑到公用桌面
Import-Module ((Split-Path $PSCommandPath) + "\copy-shortcut.ps1")
copy-shortcut

#桌面環境設定,啟用桌面圖示
Import-Module ((Split-Path $PSCommandPath) + "\Enable-DesktopIcons.ps1")
Enable-DesktopIcons 

#設定螢幕保護程式
Import-Module ((Split-Path $PSCommandPath) + "\Set-ScreenSaver.ps1")
Set-ScreenSaver

#啟用NumLock
import-module ((split-path $PSCommandPath) + "\check-numlock.ps1")
#check-numlock

#檢查VNC設定檔和服務
Import-Module ((Split-Path $PSCommandPath) + "\check-VncSetting.ps1")
Check-VncSetting
Check-VncService

#檢查firewall 有無開啟5900 5800 埠.
Import-Module ((Split-Path $PSCommandPath) + "\check-firewallport.ps1")
check-Firewallport

#檢查firewall 有無開啟VNC程式通過.
Import-Module ((Split-Path $PSCommandPath) + "\check-firewallsettings.ps1")
check-FirewallSettings

#執行3個環境設定檔 
Import-Module ((Split-Path $PSCommandPath) + "\check-VGHTCenv.ps1")
check-VGHTCenv

#設定SmartIris
Import-Module ((split-path  $PSCommandPath) + "\check-SmartIris.ps1")
#Check-SmartIris


#啟用微軟倉頡輸入法.
Import-Module ((Split-Path $PSCommandPath) + "\Enable-ChangJieinput.ps1")
enable-ChangJieinput

#設定Edge開啟IE為永不, IE停用IEtoEdge元件, 清除瀏覽器預設值."
Import-Module ((Split-Path $PSCommandPath) + "\set-IEtoEdageNever.ps1")
set-IEtoEdageNever

#設定IE,Edage,Chrome 預設開啟首頁為 "https://eip.vghtc.gov.tw"
Import-Module ((Split-Path $PSCommandPath) + "\set-HomePage.ps1")
set-HomePage

#檢查雲端藥歷的設定
Import-Module ((Split-Path $PSCommandPath) + "\check-cloudMED.ps1")
check-cloudMED


#關閉win11升級提示
Import-Module ((Split-Path $PSCommandPath) + "\disable-win11upgrade.ps1")
disable-win11upgrade


#修補
Import-Module ((Split-Path $PSCommandPath) + "\check-patch.ps1")
#check-patch

#清理windows 暫存
Import-Module ((Split-Path $PSCommandPath) + "\Clear-WindowsJunk.ps1")
Clear-WindowsJunk



#結束, 結束記錄
Stop-Transcript

pause