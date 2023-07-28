
param($runadmin)

function Set-ScreenSaver {
    write-output "�]�w�ù��O�@�{��"
  
    # �]�w�S�w���ù��O�@�{���ɮ׸��|
    $screenSaverFilePath = "c:\screensaver.scr"

    # �]�w�ù��O�@�{��
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaveActive -Value 1

    # �]�w�ù��O�@�{�����ݮɶ��]�H�����^
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaveTimeOut -Value 900

    # �]�w�ù��O�@�{�����K�X�O�@���A�]0��ܸT�ΡA1��ܱҥΡ^
    #Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name ScreenSaverIsSecure -Value 1

    # �]�w�S�w���ù��O�@�{���ɮ׸��|
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name SCRNSAVE.EXE -Value $screenSaverFilePath

}

#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q?�J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    set-ScreenSaver
    pause
}
