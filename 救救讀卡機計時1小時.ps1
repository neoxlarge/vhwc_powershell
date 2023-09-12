#從裝置管理員中移除?藏的設備
#
#只會移除ScmarCard ,SmartCardReader和SmartCardFilter這3個和讀卡機相關的?藏的設備.
#powershell V2無法執行, get-pnpdevice是V5的語法.
#
#win32_pnpentity中不會列出?藏的設備, 無法用win32_pnpentity來作.
#2020819, 

param($runadmin)


function remove-HiddenDevice {

  $Username = "vhcy\vhwcmis"
  $Password = "Mis20190610"
  $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
  $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

  # 部分舊win10系統旳pnputil.exe沒有remove-device功能, 目前手上己更新的22H2中的版本為10.0.19041.3155, 以此為準. 
  # 比這版舊的都更新到這版.

  $pnputil_version = Get-ItemPropertyValue -Path "$env:windir\system32\pnputil.exe" -name VersionInfo

  if ($pnputil_version.FileVersion -lt [version]"10.0.19041.3324") {
    #版本低於10.1.19041.3324, 在底下的路徑中找到復制符合的版本, 復制到c:\windows\system32
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
  Write-Host "救救讀卡機"
  Write-Output "$($pnputil_version.Filename) : $($pnputil_version.FileVersion)"

  # todo:  win7要用devcon.exe, win10也有但要另安裝windows10 WDK
  # Win7的要再測看看.
  #devcon.exe download: https://superuser.com/questions/1002950/quick-method-to-install-devcon-exe

  do {
    $dev = Get-PnpDevice | Where-Object -FilterScript { $_.Present -eq $false -and $_.Class -in ('SmartCard', 'SmartCardReader', 'SmartCardFilter') }
    Write-Output " Time: $(Get-Date) , Find Devie couts : $($dev.count)"

    if ($dev.count -ne 0) {

      foreach ($d in $dev) {
        
        if ($check_admin) {
          #登入者有管理者權限
          Write-Output "(Admin)刪除設備: $($d.FriendlyName)"
          Start-Process -FilePath "pnputil.exe" -ArgumentList "/remove-device $($d.instanceID)" -Wait -NoNewWindow
        }
        else {
          Write-Output "(User)刪除設備: $($d.FriendlyName)"
          $result = Start-Process -FilePath "pnputil.exe" -ArgumentList "/remove-device $($d.instanceID)" -Credential $credential -PassThru -NoNewWindow #這行用-wait會出現權限不足, 以下行替代.
          $result.WaitForExit()
                    
        }
      }
    }
    
    Start-Sleep -Seconds 3600
  } until ( $false )


}

  
#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {

  #檢查是否管理員
  $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

  if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
  }

  remove-HiddenDevice
  pause
}