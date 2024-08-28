#�q�˸m�޲z�����������ê��]��
#
#�u�|����ScmarCard ,SmartCardReader�MSmartCardFilter, USB, Keyboard��USB�MŪ�d�����������ê��]��.
#powershell V2�L�k����, get-pnpdevice�OV5���y�k.
#
#win32_pnpentity�����|�C�X���ê��]��, �L�k��win32_pnpentity�ӧ@.
#
#�D�n�Hpnputil.exe�Ӱ���,�ҥHwin7�����, Win10�Y��pnptuil�����L�¨S��/remove-device�Ѽ�, �]�������.
# FIXME: can not run on remore mode.


#��D���x��QuickEdit�����i�H���ϥΪ̤��p�߫���powershell console, �𦨼Ȱ�.
Set-ItemProperty  -Path "HKCU:\Console" -Name QuickEdit -Value 0

#�ˬd�O�_�޲z��
$check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")


function Compare-Versions {
  <#���2�Ӫ���, $version1 �j�󵥩� $version2 �^��$Ture #>
  param (
    [Parameter(Mandatory = $true)]
    [string]$Version1, # �Ĥ@�Ӫ���

    [Parameter(Mandatory = $true)]
    [string]$Version2     # �ĤG�Ӫ���
  )

  # �N������������}�C�A�H�K�v�Ӥ���U�ӳ���
  $version1Array = $Version1.Split('.')
  $version2Array = $Version2.Split('.')

  # �ϥ� foreach �j��M���C�ӳ����i����
  foreach ($i in 0..$version1Array.Count) {
    if ([int]$version1Array[$i] -gt [int]$version2Array[$i]) {
      return $true    # ��^ $true ��ܲĤ@�Ӫ������j��ĤG�Ӫ�����
    }
    elseif ([int]$version1Array[$i] -lt [int]$version2Array[$i]) {
      return $false   # ��^ $false ��ܲĤ@�Ӫ������p��ĤG�Ӫ�����
    }
    else {
      # �p�G��e�����۵��A�h�~�����U�@�ӳ���
      continue
    }
  }

  # �p�G�����ۦP�A�h��ܪ������ۦP
  return $true    # ��^ $true ��ܨ�Ӫ������ۦP
}


function remove-HiddenDevice {

  $Username = ".\USER"
  $Password = "Us2791072"
  $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
  $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)


  # ������win10�t����pnputil.exe�S��remove-device�\��, �ثe��W�v��s��22H2����������10.0.19041.3155, �H������. 
  # ��o���ª�����s��o��.

  #pnputil.exe ���䥻���̪�, �p�G�S���A�_������W����$env:temp
  $pnputil_path = $null

  # https://github.com/MScholtes/PS2EXE#script-variables
  if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
  { $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
  else {
    $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath) { $ScriptPath = "." } 
  }


  $pnputil_files = @(
    "C:\Windows\system32\pnputil.exe",
    "$((get-item -path $ScriptPath).DirectoryName)\pnputil.exe"
  )
  foreach ($p in $pnputil_files) {
    $p_version = Get-ItemPropertyValue -Path $p -Name VersionInfo -ErrorAction SilentlyContinue
    if ($p_version -ne $null) {
      if ((Compare-Versions -Version1 $p_version.ProductVersion -Version2 "10.0.19041.3155")) {
        $pnputil_path = $p
        break
      }
    }
  }

  if (!$pnputil_path) {
    $pnputil_files = @(
      "\\172.20.5.185\powershell\vhwc_powershell\pnputil.exe",
      "\\172.20.1.122\share\software\00newpc\vhwc_powershell\pnputil.exe"
    )

    foreach ($p in $pnputil_files) {
      $p_version = Get-ItemPropertyValue -Path $p -Name VersionInfo
      if ((Compare-Versions -Version1 $p_version.ProductVersion -Version2 "10.0.19041.3155")) {
        Copy-Item -Path $p -Destination $env:TEMP -Force -Verbose
        $pnputil_path = "$($env:temp)\pnputil.exe"
        break
      }
    }

  }

  $pnputil_version = Get-ItemPropertyValue -Path $pnputil_path -name VersionInfo
  
  Write-Output "$($pnputil_version.Filename) : $($pnputil_version.FileVersion)"

  $dev = Get-PnpDevice | Where-Object -FilterScript { $_.Present -eq $false -and $_.Class -in ('SmartCard', 'SmartCardReader', 'SmartCardFilter', 'USB', 'HIDClass', 'Keyboard','mouse','monitor','volume') }

  Write-Output "Find devices amount: $($dev.count)"

  # todo:  win7�n��devcon.exe, win10�]�����n�t�w��windows10 WDK
  # Win7���n�A���ݬ�.

  foreach ($d in $dev) {
    if ($check_admin) {
      #�n�J�̦��޲z���v��
      Write-Output "(Admin)Dedete device: $($d.FriendlyName)"
      Start-Process -FilePath $pnputil_path -ArgumentList "/remove-device $($d.instanceID)" -Wait -NoNewWindow
    }
    else {
      Write-Output "(User)Dedete device: $($d.FriendlyName)"
      $result = Start-Process -FilePath $pnputil_path -ArgumentList "/remove-device $($d.instanceID)" -Credential $credential -PassThru -NoNewWindow #�o���-wait�|�X�{�v������, �H�U����N.
      $result.WaitForExit()
    }

  }

}

#mian

remove-HiddenDevice

Start-Sleep -Seconds 10