#�q�˸m�޲z��������?�ê��]��
#
#�u�|����ScmarCard ,SmartCardReader�MSmartCardFilter�o3�өMŪ�d��������?�ê��]��.
#powershell V2�L�k����, get-pnpdevice�OV5���y�k.
#
#win32_pnpentity�����|�C�X?�ê��]��, �L�k��win32_pnpentity�ӧ@.
#2020819, 

param($runadmin)


function remove-HiddenDevice {

  $Username = "vhcy\vhwcmis"
  $Password = "Mis20190610"
  $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
  $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

 # $dev =  Get-PnpDevice | Where-Object -FilterScript {$_.Present -eq $false -and $_.Class -in ('SmartCard','SmartCardReader','SmartCardFilter')}

  # ������win10�t����pnputil.exe�S��remove-device, �ҥHcopy�F�@��, �A��os�̪���@�U�s��.
  $pnputil_paths = "$((get-item -path $PSCommandPath).DirectoryName)\pnputil.exe",
            "$($env:USERPROFILE)\desktop\vhwc_powershell\pnputil.exe",
            "D:\mis\vhwc_powershell\pnputil.exe"
  foreach ($p in $pnputil_paths) {
    if (Test-Path -Path $p)
    {
      $pnputil = $p
      break
    }
  }

  $pnputil_os = "C:\Windows\system32\pnputil.exe"
  $result = (Get-ItemPropertyValue -Path $pnputil -name VersionInfo).productVersion -lt (Get-ItemPropertyValue -Path $pnputil_os -name VersionInfo).productVersion
  #Write-Host (Get-ItemPropertyValue -Path $pnputil -name VersionInfo).productVersion
  #Write-Host (Get-ItemPropertyValue -Path $pnputil_os -name VersionInfo).productVersion
  
  if ($result) {$pnputil = $pnputil_os}
  Write-Output "pnputil path: $pnputil"

  # todo:  win7�n��devcon.exe, win10�]�����n�t�w��windows10 WDK
  # Win7���n�A���ݬ�.
  #devcon.exe download: https://superuser.com/questions/1002950/quick-method-to-install-devcon-exe

  do {
    $dev =  Get-PnpDevice | Where-Object -FilterScript {$_.Present -eq $false -and $_.Class -in ('SmartCard','SmartCardReader','SmartCardFilter')}

    if ($dev.count -ne 0) {
      foreach ($d in $dev) {
        
        if ($check_admin) {
          #�n�J�̦��޲z���v��
          Write-Output "(Admin)�R���]��: $($d.FriendlyName)"
          Start-Process -FilePath $pnputil -ArgumentList "/remove-device $($d.instanceID)" -Wait -NoNewWindow
        } else {
          Write-Output "(User)�R���]��: $($d.FriendlyName)"
          #Start-Process -FilePath $pnputil -ArgumentList "/remove-device $($d.instanceID)" -Credential $credential -Wait -NoNewWindow
          #�p�G�S�v���N�ݥ�invoke-command
          Invoke-Command -ScriptBlock {
            Param (
             [string]$PnpUtilPath,
             [string]$DeviceInstanceID)

            Start-Process -FilePath $pnputil -ArgumentList "/remove-device $DeviceInstanceID" -Wait -NoNewWindow  
          } -Credential $credential -ComputerName "localhost" -ArgumentList $pnputil,$($d.instanceID)
        }
      }


    }
    
    Start-Sleep -Seconds 3600
  } until ( $false )


}

  
#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q�פJ�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    remove-HiddenDevice
    pause
}