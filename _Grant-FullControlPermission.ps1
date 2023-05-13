#�ק�T���s�������w�]����, IE, Edge, chrome.
#20230329, ����@��ϥΪ̵L�k�g�JHKCU:\SOFTWARE\Policies, �令�Ѻ޲z�̼g��HKLM:\SOFTWARE\Policies
param($runadmin)

function Grant-FullControlPermission {
    <#
    ��ƦW�٬� Grant-FullControlPermission�A����ӰѼơGFolders�]�]�t�n�¤����������v������Ƨ��M��^�M UserName�]�n���v�����ϥΪ̦W�١^�C
    �b��Ƥ����A�ϥ� foreach �j��M����Ƨ��M��A��C�Ӹ�Ƨ�����ۦP���ާ@�G���o ACL�B�إߦs���W�h�B�s�W�W�h�� ACL�A�M��N�ק�᪺ ACL �M�Φܸ�Ƨ��C
    �̫�A�z�i�H�I�s�Ө�ơA��J��Ƨ��M��M�ϥΪ̦W�١A�H�¤����w���ϥΪ̧��������v���C
    #>

    $folders = "c:\2100", "C:\oracle", "c:\mis", "C:\cloudMED", "C:\ICCARD_HIS", "C:\IDMSClient45", "C:\NHI", "C:\TEDPC", "C:\VGHTC","C:\VghtcLogo", "C:\vhgp"
    $userName = "User"

    foreach ($folderPath in $Folders) {
        # ���o��Ƨ��� ACL
        $acl = Get-Acl -Path $folderPath
        
        # �إߤ@�ӷs���s���W�h�A�¤����w�ϥΪ̧��������v��
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($UserName, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
        
        # �N�s���W�h�s�W�� ACL
        $acl.SetAccessRule($rule)
        
        # �N�ק�᪺ ACL �M�Φܸ�Ƨ�
        Set-Acl -Path $folderPath -AclObject $acl
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
