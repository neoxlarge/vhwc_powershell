#�]����v�n�D���ୱ�ѪR�פ��P,�s�a�n�ק�Pacs���ù��t�m.
#�|���n�D�r��j, �ҥH�ѪR�׽էC
#�]��v�n�D�ѪR�׽հ�, �ҥH���~�i��.

#AD�W�n�إ�2�ӱb��, �@�ӬO�̨θѪR��:OPD-C09HighResolution,�@�ӬO�j�r��:OPD-PC09BigFont, �K�X�]�@��
#�q���W�n�}powershel(admin mode), ������O,�ܧ�w�]����W�h Set-ExecutionPolicy -ExecutionPolicy RemoteSigned, ���t�Υi�H����.ps1��

#�ǳ�2��pacs������]�w��,IRIS_Sys_Highresolution.ini �M IRIS_Sys_BigFont.ini, �̨Ϥ���л\���U�]�w��.
#C:\TEDPC\SmartIris\SysIni\IRIS_Sys.ini

$ini_list = @{
    highresolution = "C:\TEDPC\SmartIris\SysIni\IRIS_Sys_Highresolution.ini";
    bigfont = "C:\TEDPC\SmartIris\SysIni\IRIS_Sys_BigFont.ini";
    original = "C:\TEDPC\SmartIris\SysIni\IRIS_Sys.ini"
}

switch ($env:USERNAME) {
    OPD-C09HighResolution { 
        Copy-Item -Path $ini_list.highresolution -Destination $ini_list.original -Force
     }
    OPD-PC09BigFont { 
        copy-item -Path $ini_list.bigfont -Destination $ini_list.original -Force
    }
    Default {}
}
