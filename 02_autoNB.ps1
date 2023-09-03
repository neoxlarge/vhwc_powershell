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

#==�w��================================================================== 

#�ˬd�αҥ�SMBv1/CIFS�\��. ���\�ॻ���ӭ��}��,�����������}��,�������n��˧�.
import-module ((split-path $PSCommandPath) + "\Enable-Win10OPTFeature.ps1")
Enable-SMBv1
Enable-NetFx3

# �_��VGHTC����|�M�ε{��.
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

## �w��CMS_CGServiSignAdapter
### �̤��n�D,�w�˫e���������r�n��, �ҥH�񨾬r���w��
Import-Module ((Split-Path $PSCommandPath) + "\install-CMS.ps1")
install-CMS


## �w�� HCAServiSign
Import-Module ((Split-Path $PSCommandPath) + "\install-HCA.ps1")
install-HCA

# �w��Hicos
Import-Module ((Split-Path $PSCommandPath) + "\install-HiCOS.ps1")
install-Hicos

# �w��IDC
Import-Module ((Split-Path $PSCommandPath) + "\install-IDC.ps1")
install-IDC

#�w��VNC
Import-Module ((Split-Path $PSCommandPath) + "\install-VNC.ps1")
install-vnc

#install java
Import-Module ((Split-Path $PSCommandPath) + "\install-Java.ps1")
install-java
set-Java_env

#�w��2100
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

# �w�˶��ݦw�����󰷫O�dŪ�d������(PCSC)
Import-Module ((Split-Path $PSCommandPath) + "\install-PCSC.ps1")
install-PCSC

#�w�˵������O�d
Import-Module ((Split-Path $PSCommandPath) + "\install-virtualnhc.ps1")
install-virtualnhc

#�w�ˮw���f�Ľ]�d�ߨt��
Import-Module ((Split-Path $PSCommandPath) + "\install-cdcalert.ps1")
install-cdcalert


# �w��sslvpn
Import-Module ((Split-Path $PSCommandPath) + "\install-sslvpn.ps1")
install-SMAConnectAgent
install-NX

#�w��anydesk
 Import-Module ((Split-Path $PSCommandPath) + "\install-AnyDesk.ps1")
 install-AnyDesk

# �w��Winnexus
Import-Module ((Split-Path $PSCommandPath) + "\install-Winnexus.ps1")

#install-winNexus    

# �w�˨��r Trend Micro Apex One Security Agent
Import-Module ((Split-Path $PSCommandPath) + "\install-AntiVir.ps1")

install-antivir

# ���������n��win10 �{��
Import-Module ((Split-Path $PSCommandPath) + "\remove-apps.ps1")
remove-apps


####�t�γ]�w�s######################################################################################

#�ҥ�powershell���ݺ޲z
import-module ((split-path $PSCommandPath) + "\Check-EnablePSRemoting.ps1")
#check-enablepsremoting

##�ק��ƪ��v��
import-module ((split-path $PSCommandPath) + "\grant-FullControlPermission.ps1")
Grant-FullControlPermission

#�ܧ�q���p�e
import-module ((split-path $PSCommandPath) + "\disable-sleep.ps1")
#disable-sleep

#����UAC
Import-Module ((Split-Path $PSCommandPath) + "\Disable-UAC.ps1")
Disable-UAC

#�_��|�줽�ήୱ
Import-Module ((Split-Path $PSCommandPath) + "\copy-shortcut.ps1")
copy-shortcut

#�ୱ���ҳ]�w,�ҥήୱ�ϥ�
Import-Module ((Split-Path $PSCommandPath) + "\Enable-DesktopIcons.ps1")
Enable-DesktopIcons 

#�]�w�ù��O�@�{��
Import-Module ((Split-Path $PSCommandPath) + "\Set-ScreenSaver.ps1")
Set-ScreenSaver

#�ҥ�NumLock
import-module ((split-path $PSCommandPath) + "\check-numlock.ps1")
#check-numlock

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

#�]�wSmartIris
Import-Module ((split-path  $PSCommandPath) + "\check-SmartIris.ps1")
#Check-SmartIris


#�ҥηL�n�ܾe��J�k.
Import-Module ((Split-Path $PSCommandPath) + "\Enable-ChangJieinput.ps1")
enable-ChangJieinput

#�]�wEdge�}��IE���ä�, IE����IEtoEdge����, �M���s�����w�]��."
Import-Module ((Split-Path $PSCommandPath) + "\set-IEtoEdageNever.ps1")
set-IEtoEdageNever

#�]�wIE,Edage,Chrome �w�]�}�ҭ����� "https://eip.vghtc.gov.tw"
Import-Module ((Split-Path $PSCommandPath) + "\set-HomePage.ps1")
set-HomePage

#�ˬd�����ľ����]�w
Import-Module ((Split-Path $PSCommandPath) + "\check-cloudMED.ps1")
check-cloudMED


#����win11�ɯŴ���
Import-Module ((Split-Path $PSCommandPath) + "\disable-win11upgrade.ps1")
disable-win11upgrade


#�׸�
Import-Module ((Split-Path $PSCommandPath) + "\check-patch.ps1")
#check-patch

#�M�zwindows �Ȧs
Import-Module ((Split-Path $PSCommandPath) + "\Clear-WindowsJunk.ps1")
Clear-WindowsJunk



#����, �����O��
Stop-Transcript

pause