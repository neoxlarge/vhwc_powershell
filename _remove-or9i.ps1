
param($runadmin)

function remove-or9i {

    $path = "HKLM:\SOFTWARE\WOW6432Node\ORACLE", 
            "C:\oracle",
            "C:\Program Files (x86)\Oracle"
            "C:\Program Files\Oracle"

    foreach ($p in $path) {
        if (Test-Path -Path $p) {
            Remove-Item -Path $p -Recurse -Force
        }


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

    remove-or9i
    pause
}
