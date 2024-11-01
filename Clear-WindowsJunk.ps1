#�M�zWindows, chrome, edge, ie���Ȧs��
param($runadmin)


function Clear-WindowsJunk {
  # �M���Τ�Ȧs�� 
  Write-Output "�M���Τ�Ȧs��:"
  
  $user_folders = Get-ChildItem "C:\Users"
  foreach ($user in $user_folders) {

    Write-OutPut "Users\$user"
    Remove-Item -Path "c:\Users\$user\AppData\Local\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "c:\Users\$user\AppData\Local\Microsoft\Windows\INetCache\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "c:\Users\$user\AppData\Local\Microsoft\Windows\INetCookies\*" -Recurse -Force -ErrorAction SilentlyContinue
    
  }

 
  Write-Output "�M���t�μȦs��$($env:windir)\Temp\*"
  Remove-Item -Path "$env:windir\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue

 
  Write-Output "�M��Windows ��s�����ͦ����Ȧs��"
  Remove-Item -Path "$env:windir\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue

  
  Write-Output "�M��windows �֨���Ƨ�"
  Remove-Item -Path "C:\Windows\Prefetch\*" -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -Path "C:\Windows\SystemTemp\*" -Recurse -Force -ErrorAction SilentlyContinue
  
    
  Write-OutPut " �M��Google Chrome�Ȧs�ɩMCookies"
  try {
      Remove-Item -Path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -ErrorAction  SilentlyContinue
      Remove-Item -Path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Network\cookies" -Recurse -Force -ErrorAction SilentlyContinue 
  }
  catch [System.Management.Automation.ItemNotFoundException] {
      Write-Warning "�M���Ȧs�ɥi�ॢ��:"
      Write-Warning $Error[0].Exception.Message
  }

    
    Write-OutPut " �M��Microsoft Edge�Ȧs�ɩMCookies"
  try {
    Remove-Item -Path "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\Cookies\*" -Recurse -Force -ErrorAction SilentlyContinue
  }
    catch [System.Management.Automation.ItemNotFoundException]{
      Write-Warning "�M���Ȧs�ɥi�ॢ��:"
      Write-Warning $Error[0].Exception.Message
  }
  

  <# �M��Internet Explorer�Ȧs�ɩMCookies
      Delete Temporary Internet Files:
      RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8

      Delete Cookies:
      RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2

      Delete History:
      RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1

      Delete Form Data:
      RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16

      Delete Passwords:
      RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32

      Delete All:
      RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255

      Delete All + files and settings stored by Add-ons:
      RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351
  #>

  Write-OutPut "�M��Internet Explorer�Ȧs�ɩMCookies"
  Start-Process -FilePath "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess 8" -Wait
  Start-Process -FilePath "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess 2" -Wait

  # �M�Ŧ^���� , 
  # 20240403, ���n�M�Ŧ^����, �i�঳�ϥΪ̼ȯd�����.
  #Write-OutPut "�M�Ŧ^����"
  #Clear-RecycleBin -Force -ErrorAction SilentlyContinue

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