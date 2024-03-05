
param($runadmin)


function disable-sleep {
    Write-Output "�ܧ�q���p�e:"

    write-output "�]�w�q���p��:����"
    powercfg /setactive "SCHEME_BALANCED"

    Write-Output "�w��-�����w�Ыe���ɶ�:0"
    powercfg /change disk-timeout-ac 0

    write-output  "������ܾ�:15��"
    powercfg /change monitor-timeout-ac 15
    
    write-output "���q���ίv:�ä�"
    powercfg /change standby-timeout-ac 0

    Write-Output "�����V�X���ίv:����"
    powercfg /hibernate off

}

#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q�פJ�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    disable-sleep
    pause
}