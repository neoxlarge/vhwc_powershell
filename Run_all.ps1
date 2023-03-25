# 自動安裝和設定
param($runadmin)
$run_main = $true

# 以管理員權限執行, Elevate Powershell to Admin
$check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (!$check_admin -and !$runadmin) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -runadmin 1" -Verb RunAs; exit
}

#啟用記錄
Start-Transcript -Path "d:\mis\$(Get-Date -Format 'yyyy-MM-dd-hh-mm').log" -Append

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


#==設定===================================================================

#啟用powershell遠端管理
import-module ((split-path $PSCommandPath) + "\Check-EnablePSRemoting.ps1")

check-enablepsremoting


#檢查及啟用SMBv1/CIFS功能.
import-module ((split-path $PSCommandPath) + "\Check-smbcifs.ps1")

check-smbcifs


#啟用NumLock
import-module ((split-path $PSCommandPath) + "\check-numlock.ps1")

check-numlock


#復制捷徑到公用桌面
Import-Module ((Split-Path $PSCommandPath) + "\copy-shortcut.ps1")

copy-shortcut


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


#啟用微軟倉頡輸入法.

Import-Module ((Split-Path $PSCommandPath) + "\Enable-ChangJieinput.ps1")

enable-ChangJieinput


#公文系統?境設定. 

Import-Module ((Split-Path $PSCommandPath) + "\Check-2100env.ps1")

Check-2100env


#桌面環境設定,啟用桌面圖示
Import-Module ((Split-Path $PSCommandPath) + "\Enable-DesktopIcons.ps1")

Enable-DesktopIcons 

#==安裝================================================================== 


## 安裝CMS_CGServiSignAdapter
### 依文件要求,安裝前應關閉防毒軟體, 所以比防毒先安裝
Import-Module ((Split-Path $PSCommandPath) + "\install-CMS.ps1")

install-CMS


## 安裝 HCAServiSign
Import-Module ((Split-Path $PSCommandPath) + "\install-HCA.ps1")

install-HCA


# 安裝雲端安控元件健保卡讀卡機控制(PCSC)
Import-Module ((Split-Path $PSCommandPath) + "\install-PCSC.ps1")

install-PCSC


# 安裝Winnexus
Import-Module ((Split-Path $PSCommandPath) + "\install-Winnexus.ps1")

install-winNexus    


# 安裝防毒 Trend Micro Apex One Security Agent
Import-Module ((Split-Path $PSCommandPath) + "\install-OfficeScan.ps1")

install-OfficeScan


#結束, 結束記錄
Stop-Transcript

pause