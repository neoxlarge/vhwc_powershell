#��IE�]���w�]��.
#1. ����EDGE����IE�Ҧ�.
#2. ����IE����IEtoEdga����
#3. ��IE�]���w�]��, �o�ʧ@�L�k�����ק�registry, �]���L�k��Xhash.
#   �ݭngroup policy���覡?��.

param($runadmin)

set-IEasDefault {

    Write-Output "�bEdge���HIE�}�Һ����]���ä�"
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Edge\IEToEdge" -Name "RedirectionMode" -Value 0 -Force

    Write-Output "����IE����IEtoEdga����"
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Ext\CLSID" -Name "{1FD49718-1D00-4B19-AF5F-070AF6D5D54C}" -Value 0 -Force

    

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
