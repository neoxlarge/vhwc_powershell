# �۰ʦw�˩M�]�w
param($runadmin)
$run_main = $true

# �H�޲z���v������, Elevate Powershell to Admin
$check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (!$check_admin -and !$runadmin) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -runadmin 1" -Verb RunAs; exit
}

#�ҥΰO��
Start-Transcript -Path "d:\mis\$(Get-Date -Format 'yyyy-MM-dd-hh-mm').log" -Append

if ($check_admin) {
    Write-Output "Powershell run as Admin mode."
}
else {
    Write-Output "Powershell run as User mode."
}

#�ˬdpowershell����, ���Ӱ���5.1, Win7�w����2.0, Win10�w����5.1, Win7�ݭn�w��5.1
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "Powershell �������ӭn5.1�H�W, �������." -ErrorAction Stop
}


#==�]�w===================================================================

#�ҥ�powershell���ݺ޲z
import-module ((split-path $PSCommandPath) + "\Check-EnablePSRemoting.ps1")

check-enablepsremoting


#�ˬd�αҥ�SMBv1/CIFS�\��.
import-module ((split-path $PSCommandPath) + "\Check-smbcifs.ps1")

check-smbcifs


#�ҥ�NumLock
import-module ((split-path $PSCommandPath) + "\check-numlock.ps1")

check-numlock


#�_��|�줽�ήୱ
Import-Module ((Split-Path $PSCommandPath) + "\copy-shortcut.ps1")

copy-shortcut


#�ˬdVNC�]�w�ɩM�A��
Import-Module ((Split-Path $PSCommandPath) + "\check-VncSetting.ps1")

Check-VncSetting
Check-VncService

#�ˬdfirewall ���L�}��5900 5800 ��.
Import-Module ((Split-Path $PSCommandPath) + "\check-firewallport.ps1")

check-Firewallport

#�ˬdfirewall ���L�}��VNC�{���q�L.
Import-Module ((Split-Path $PSCommandPath) + "\check-firewallsettings.ps1")

check-FirewallSettings


#����3�����ҳ]�w�� 
Import-Module ((Split-Path $PSCommandPath) + "\check-VGHTCenv.ps1")

check-VGHTCenv


#�ҥηL�n�ܾe��J�k.

Import-Module ((Split-Path $PSCommandPath) + "\Enable-ChangJieinput.ps1")

enable-ChangJieinput


#����t��?�ҳ]�w. 

Import-Module ((Split-Path $PSCommandPath) + "\Check-2100env.ps1")

Check-2100env


#�ୱ���ҳ]�w,�ҥήୱ�ϥ�
Import-Module ((Split-Path $PSCommandPath) + "\Enable-DesktopIcons.ps1")

Enable-DesktopIcons 

#==�w��================================================================== 


## �w��CMS_CGServiSignAdapter
### �̤��n�D,�w�˫e���������r�n��, �ҥH�񨾬r���w��
Import-Module ((Split-Path $PSCommandPath) + "\install-CMS.ps1")

install-CMS


## �w�� HCAServiSign
Import-Module ((Split-Path $PSCommandPath) + "\install-HCA.ps1")

install-HCA


# �w�˶��ݦw�����󰷫O�dŪ�d������(PCSC)
Import-Module ((Split-Path $PSCommandPath) + "\install-PCSC.ps1")

install-PCSC


# �w��Winnexus
Import-Module ((Split-Path $PSCommandPath) + "\install-Winnexus.ps1")

install-winNexus    


# �w�˨��r Trend Micro Apex One Security Agent
Import-Module ((Split-Path $PSCommandPath) + "\install-OfficeScan.ps1")

install-OfficeScan


#����, �����O��
Stop-Transcript

pause