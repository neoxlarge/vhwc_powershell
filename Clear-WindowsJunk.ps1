#�M�zWindows���Ȧs�M�^����
param($runadmin)


function Clear-WindowsJunk {
  # �M���Τ�Ȧs�� 
  Write-Output "�M���Τ�Ȧs��$env:LOCALAPPDATA\Temp\*"
  Remove-Item -Path "$env:LOCALAPPDATA\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue

  Write-Output "�M���t�μȦs��env:windir\Temp\*"
  # �M���t�μȦs�� 
  Remove-Item -Path "$env:windir\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue

  # �M�Ŧ^���� 
  Write-OutPut "�M�Ŧ^����"
  Clear-RecycleBin -Force -ErrorAction SilentlyContinue

  # �M���ƥ�d�ݾ���x 
  #WEVTUtil.exe cl Application
  #WEVTUtil.exe cl System
  #WEVTUtil.exe cl Security
}

  
#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q�פJ�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Clear-WindowsJunk
    pause
}