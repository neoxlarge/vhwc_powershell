#�ק�]�Ʊ����v, �i��AD���ϥΪ̧��L����]�w.

param($runadmin)

function Grant-DeviceOwner {
   

    if ($check_admin) {
        # �w�q�����]�ƾ֦��̸s�զW�٩M����ϥΪ̸s�զW��
        $localGroup = "Device Owner"
        $domainGroup = "vhcy\Domain Users"  

        # ���o�����]�ƾ֦��̸s�ժ���
        $localGroupObject = [ADSI]"WinNT://$env:COMPUTERNAME/$localGroup,group"

        # ���o����ϥΪ̸s�ժ���
        $domainGroupObject = [ADSI]"WinNT://$domainGroup,group"

        # �N����ϥΪ̸s�շs�W�ܥ����]�ƾ֦��̸s�դ�
        $localGroupObject.Add($domainGroupObject.Path)

        Write-Host "����ϥΪ̸s�դw�s�W�ܳ]�ƾ֦��̸s�դ��C"
    }
    else {
        Write-Warning "�S���t�κ޲z���v��,�L�k�}�Ҹ��?�v��,�ХH�t�κ޲z���������s����."
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

    Grant-FullControlPermission
    pause
}
