# �۰ʦw�˩M�]�w
param($run_admin)
$run_main = $true


# �H�޲z���v������, Elevate Powershell to Admin
$check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (!$check_admin -and !$run_admin) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -run_admin 1" -Verb RunAs; exit
}


#�ҥΰO��
#�slog �ɪ����?
if ((Test-Path -Path "d:\")) {
    $log_folder = "d:\mis"
} else {$log_folder = "c:\mis"}
Start-Transcript -Path "$log_folder\$(Get-Date -Format 'yyyy-MM-dd-hh-mm').log" -Append

#�O��powershell�O�Τ����v�����檺
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
import-module ((split-path $PSCommandPath) + "\enable-Win10OPTFeature.ps1")

enable-smbv1
enable-NetFx3


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

#�]�wIE,Edage,Chrome �w�]�}�ҭ����� "https://eip.vghtc.gov.tw"
Import-Module ((Split-Path $PSCommandPath) + "\set-HomePage.ps1")

set-HomePage


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
#Import-Module ((Split-Path $PSCommandPath) + "\install-Winnexus.ps1")

#install-winNexus    


# �w�˨��r Trend Micro Apex One Security Agent
#Import-Module ((Split-Path $PSCommandPath) + "\install-AntiVir.ps1")

#install-antivir

# ���������n��win10 �{��
Import-Module ((Split-Path $PSCommandPath) + "\remove-apps.ps1")

remove-apps

#����, �����O��
Stop-Transcript

pause