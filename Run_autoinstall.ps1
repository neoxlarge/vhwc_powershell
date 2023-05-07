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

#啟用微軟倉頡輸入法.

Import-Module ((Split-Path $PSCommandPath) + "\Enable-ChangJieinput.ps1")

enable-ChangJieinput

#桌面環境設定,啟用桌面圖示
Import-Module ((Split-Path $PSCommandPath) + "\Enable-DesktopIcons.ps1")

Enable-DesktopIcons 

#==安裝================================================================== 

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

#安裝2100
Import-Module ((Split-Path $PSCommandPath) + "\install-2100.ps1")
install-2100

#install chrome
Import-Module ((Split-Path $PSCommandPath) + "\install-chrome.ps1")
install-chrome

#install smartiris
Import-Module ((Split-Path $PSCommandPath) + "\install-smartiris.ps1")
install-smartiris

#install libreoffice
Import-Module ((Split-Path $PSCommandPath) + "\install-libreoffice.ps1")
install-libreoffice

# 安裝雲端安控元件健保卡讀卡機控制(PCSC)
Import-Module ((Split-Path $PSCommandPath) + "\install-PCSC.ps1")

install-PCSC

Import-Module ((Split-Path $PSCommandPath) + "\copy-vghtc.ps1")

copy-vghtc

# 安裝Winnexus
#Import-Module ((Split-Path $PSCommandPath) + "\install-Winnexus.ps1")

#install-winNexus    


# 安裝防毒 Trend Micro Apex One Security Agent
#Import-Module ((Split-Path $PSCommandPath) + "\install-AntiVir.ps1")

#install-antivir

# 移除不必要的win10 程式
Import-Module ((Split-Path $PSCommandPath) + "\remove-apps.ps1")

remove-apps

#結束, 結束記錄
Stop-Transcript

pause