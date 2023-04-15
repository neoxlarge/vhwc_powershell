#�ୱ���ҳ]�w
#�ҥήୱ�ϥ�

param($runadmin)

#���oOS������
function Get-OSVersion {
    $os = (Get-WmiObject -Class Win32_OperatingSystem).Caption

    if ($os -like "*Windows 7*") {
        return "Windows 7"
    }
    elseif ($os -like "*Windows 10*") {
        return "Windows 10"
    }
    else {
        return "Unknown OS"
    }
}

function Enable-DesktopIcons {
    Write-host "�]�w�ୱ ."
    #Win10�MWin7���ۦP���]�w�ȧ�k

    if (!(Test-Path -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel")) {
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Force | out-null
    }

    #�q�� 
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value 0
    #�ϥΪ̤��ϥ�
    #Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{59031a47-3f72-44a7-89c5-5595fe6b30ee}" -Value 0 
    #����
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" -Value 0 
    #����x
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}" -Value 0
    #�귽�^����
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{645FF040-5081-101B-9F08-00AA002F954E}" -Value 0

    if ((Get-OSVersion) -eq "Windows 10") {
        #Windows 10 �~�ݭn����o

        # �ҥήୱ�ϥ� (Win10)
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideIcons" -Value 0
        # ���� Cortana ���s(Win10)
        Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowCortanaButton" -Value 0
        # ���� �j�M ���s(Win10)
        Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0
        # ���� �s�D�M���� ���s(Win10)
        # �Ѧ�https://admx.help/?Category=Windows_10_2016&Policy=Microsoft.Policies.Feeds::EnableFeeds
        if ($check_admin) {
            #�g�JHotkey local machine �ݭn�޲z���v��.
            $reg_path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds"
            if (Test-Path -Path $reg_path) {
                Set-ItemProperty -Path $reg_path -Name "EnableFeeds" -Value 0 -Force
            }
            else {
                New-Item -Path $reg_path -force
                New-ItemProperty -Path $reg_path -Name "EnableFeeds" -Value 0 -PropertyType DWord -force
            }
        }
    }

    #���Үୱ
    Stop-Process -Name "Explorer" -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
    
    #�����Ұʮɰ���OneDrive.
    $result = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" 
    if ($result.OneDrive -ne $null) {
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "OneDrive" -Value $null
    }
}


#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q���J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Enable-DesktopIcons
    
    pause
}