#�]�wadboe pdf���۰ʧ�s

param($runadmin)


<#
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown]
"bUpdateToSingleApp"=dword:00000000
#>

function check-pdf {

    $reg_path = "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown"
    $is_update = Get-ItemProperty -Path $reg_path -Name "bUpdateToSingleApp" -ErrorAction SilentlyContinue

    if ($is_update.bUpdateToSingleApp -eq 0) {
        Write-Output "Adobe PDF Reader �v�]�w���۰ʧ�s."
    }
    else {
        if ($check_admin) {
            Write-Output "���b�]�wAdobe PDF Reader �����۰ʷs��."
            Set-ItemProperty -Path $reg_path -Name "bUpdateToSingleApp" -Value 00000000 -type DWord -Force
        }
        else {
            Write-Warning "�S���t�κ޲z���v��,�BAdobe PDF Reader���]�w���۰ʷs,�ХH�t�κ޲z���������s����."
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

    check-pdf
    pause
}