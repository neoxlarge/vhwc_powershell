#�ק�T���s�������w�]����, IE, Edge, chrome.
#20230329, ����@��ϥΪ̵L�k�g�JHKCU:\SOFTWARE\Policies, �令�Ѻ޲z�̼g��HKLM:\SOFTWARE\Policies
param($runadmin)

function Set-HomePage {
    if ($check_admin) {
        $HomePage = "https://eip.vghtc.gov.tw"
        Write-Output "�]�wIE,Edage,Chrome �w�]�}�ҭ�����: $HomePage"
        # Modify Edge home page
        # �Ѧҳs��: https://learn.microsoft.com/en-us/deployedge/microsoft-edge-policies#restoreonstartup
        $reg_path = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\RestoreOnStartupURLs"
        $result = Test-Path -Path $reg_path
        if ($result -eq $false) {
            New-Item -Path $reg_path -force
        }
        Set-ItemProperty -Path $reg_path -Name "1" -Value $HomePage
        #Edge�i�H�]�w�h�ӡAname�ȬO�Ʀr�@���[.
        #Set-ItemProperty -Path $reg_path -Name "2" -Value "https://www.vghtc.gov.tw"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "RestoreOnStartup" -Value 4

        # Modify Chrome home page
        $reg_path = "HKLM:\Software\Policies\Google\Chrome\RestoreOnStartupURLs"
        $result = Test-Path -Path $reg_path
        if ($result -eq $false) {
            New-Item -Path $reg_path -force
        }
        Set-ItemProperty -Path $reg_path -Name "1" -Value $HomePage
        Set-ItemProperty -Path "HKLM:\Software\Policies\Google\Chrome" -Name "RestoreOnStartup" -Value 4


        # Modify Internet Explorer home page
        Set-ItemProperty -Path "HKLM:\Software\Microsoft\Internet Explorer\Main" -Name "Start Page" -Value $HomePage
        Set-ItemProperty -Path "HKLM:\Software\Microsoft\Internet Explorer\Main" -Name "Default_Page_URL" -Value $HomePage

    } else {
        Write-Warning "�S���t�κ޲z���v��,�L�k�j��]�wHomepage,�ХH�t�κ޲z���������s���աC"
    }
}




#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q?�J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    Set-HomePage
    pause
}
