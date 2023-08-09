
param($runadmin)

function remove-or9i {

    if ($check_admin) {
        $BDE = Get-WmiObject -class Win32_product | where-object -FilterScript { $_.name -eq "Borland DataBase Engine" }

        if ($BDE -ne $null) {
            $BDE.Uninstall()
        }


        $path = "HKLM:\SOFTWARE\WOW6432Node\ORACLE", 
        "C:\oracle",
        "C:\Program Files (x86)\Oracle",
        "C:\Program Files\Oracle",
        "C:\Program Files (x86)\Common Files\Borland Shared",
        "HKLM:\SOFTWARE\WOW6432Node\Borland"



        foreach ($p in $path) {
            if (Test-Path -Path $p) {
                Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

    }

}

#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q�פJ�ɤ��|����禡.
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
