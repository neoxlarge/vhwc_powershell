#從裝置管理員中移除?藏的設備
#
#只會移除ScmarCard ,SmartCardReader和SmartCardFilter, USB, Keyboard等USB和讀卡機相關的?藏的設備.
#powershell V2無法執行, get-pnpdevice是V5的語法.
#
#win32_pnpentity中不會列出?藏的設備, 無法用win32_pnpentity來作.
#
#主要以pnputil.exe來執行,所以win7不能用, Win10某些pnptuil版本過舊沒有/remove-device參數, 也不能執行.

param($runadmin)

#改主控台的QuickEdit關掉可以防使用者不小心按到powershell console, 迼成暫停.
Set-ItemProperty  -Path HKCU:\Console -Name QuickEdit -Value 0

function Compare-Versions {
  <#比對2個版本, $version1 大於等於 $version2 回傳$Ture #>
  param (
    [Parameter(Mandatory = $true)]
    [string]$Version1, # 第一個版本

    [Parameter(Mandatory = $true)]
    [string]$Version2     # 第二個版本
  )

  # 將版本號拆分成陣列，以便逐個比較各個部分
  $version1Array = $Version1.Split('.')
  $version2Array = $Version2.Split('.')

  # 使用 foreach 迴圈遍歷每個部分進行比較
  foreach ($i in 0..$version1Array.Count) {
    if ([int]$version1Array[$i] -gt [int]$version2Array[$i]) {
      return $true    # 返回 $true 表示第一個版本號大於第二個版本號
    }
    elseif ([int]$version1Array[$i] -lt [int]$version2Array[$i]) {
      return $false   # 返回 $false 表示第一個版本號小於第二個版本號
    }
    else {
      # 如果當前部分相等，則繼續比較下一個部分
      continue
    }
  }

  # 如果完全相同，則表示版本號相同
  return $true    # 返回 $true 表示兩個版本號相同
}


function remove-HiddenDevice {

  $Username = "vhcy\vhwcmis"
  $Password = "Mis20190610"
  $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
  $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)


# 部分舊win10系統旳pnputil.exe沒有remove-device功能, 目前手上己更新的22H2中的版本為10.0.19041.3155, 以此為準. 
  # 比這版舊的都更新到這版.

  #pnputil.exe 先找本機裡的, 如果沒有再復制網路上的到$env:temp
  $pnputil_path = $null

  $pnputil_files = @(
    "C:\Windows\system32\pnputil.exe",
    "$((get-item -path $PSCommandPath).DirectoryName)\pnputil.exe"
  )
  foreach ($p in $pnputil_files) {
    $p_version = Get-ItemPropertyValue -Path $p -Name VersionInfo
    if ((Compare-Versions -Version1 $p_version.ProductVersion -Version2 "10.0.19041.3155")) {
      $pnputil_path = $p
      break
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

  $dev =  Get-PnpDevice | Where-Object -FilterScript {$_.Present -eq $false -and $_.Class -in ('SmartCard','SmartCardReader','SmartCardFilter','USB','HIDClass','Keyboard')}

  # todo:  win7要用devcon.exe, win10也有但要另安裝windows10 WDK
  # Win7的要再測看看.

  foreach ($d in $dev) {
    if ($check_admin) {
      #登入者有管理者權限
      Write-Output "(Admin)刪除設備: $($d.FriendlyName)"
      Start-Process -FilePath $pnputil_path -ArgumentList "/remove-device $($d.instanceID)" -Wait -NoNewWindow
    }
    else {
      Write-Output "(User)刪除設備: $($d.FriendlyName)"
      $result = Start-Process -FilePath $pnputil_path -ArgumentList "/remove-device $($d.instanceID)" -Credential $credential -PassThru -NoNewWindow #這行用-wait會出現權限不足, 以下行替代.
      $result.WaitForExit()
  }

}

}


#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {


    Set-ItemProperty  -Path HKCU:\Console -Name QuickEdit -Value 0

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    remove-HiddenDevice

    Start-Sleep -Seconds 10

    Set-ItemProperty  -Path HKCU:\Console -Name QuickEdit -Value 1
    pause
}