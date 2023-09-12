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

  # ������win10�t����pnputil.exe�S��remove-device�\��, �ثe��W�v��s��22H2����������10.0.19041.3155, �H������. 
  # ��o���ª�����s��o��.

  $pnputil_version = Get-ItemPropertyValue -Path "$env:windir\system32\pnputil.exe" -name VersionInfo

  if ($pnputil_version.FileVersion -lt [version]"10.0.19041.3324") {
    #�����C��10.1.19041.3324, �b���U�����|�����_��ŦX������, �_���c:\windows\system32
    $pnputil_files = @( "$((get-item -path $PSCommandPath).DirectoryName)\pnputil.exe",
      "$($env:USERPROFILE)\desktop\vhwc_powershell\pnputil.exe",
      "D:\mis\vhwc_powershell\pnputil.exe",
      "\\172.20.5.185\powershell\vhwc_powershell\pnputil.exe",
      "\\172.20.1.122\share\software\00newpc\vhwc_powershell\pnputil.exe"
    )

    foreach ($f in $pnputil_files) {
      $pnputil_version = Get-ItemPropertyValue -Path $f -name VersionInfo
      if ($pnputil_version.FileVersion -ge [Version]"10.0.19041.3324") {
        Copy-Item -Path $f -Destination -Path "$env:windir\system32\" -Credential $credential -Force
        break
      }
      
    }
  }

  $pnputil_version = Get-ItemPropertyValue -Path "$env:windir\system32\pnputil.exe" -name VersionInfo
  Write-Host "�ϱ�Ū�d��"
  Write-Output "$($pnputil_version.Filename) : $($pnputil_version.FileVersion)"

  # todo:  win7�n��devcon.exe, win10�]�����n�t�w��windows10 WDK
  # Win7���n�A���ݬ�.
  #devcon.exe download: https://superuser.com/questions/1002950/quick-method-to-install-devcon-exe

  do {
    $dev = Get-PnpDevice | Where-Object -FilterScript { $_.Present -eq $false -and $_.Class -in ('SmartCard', 'SmartCardReader', 'SmartCardFilter') }
    Write-Output " Time: $(Get-Date) , Find Devie couts : $($dev.count)"

    if ($dev.count -ne 0) {

      foreach ($d in $dev) {
        
        if ($check_admin) {
          #�n�J�̦��޲z���v��
          Write-Output "(Admin)�R���]��: $($d.FriendlyName)"
          Start-Process -FilePath "pnputil.exe" -ArgumentList "/remove-device $($d.instanceID)" -Wait -NoNewWindow
        }
        else {
          Write-Output "(User)�R���]��: $($d.FriendlyName)"
          $result = Start-Process -FilePath "pnputil.exe" -ArgumentList "/remove-device $($d.instanceID)" -Credential $credential -PassThru -NoNewWindow #�o���-wait�|�X�{�v������, �H�U����N.
          $result.WaitForExit()
                    
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