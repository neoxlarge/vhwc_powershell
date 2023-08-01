#從裝置管理員中移除?藏的設備
#
#只會移除ScmarCard ,SmartCardReader和SmartCardFilter這3個和讀卡機相關的?藏的設備.
#powershell V2無法執行, get-pnpdevice是V5的語法.
#
#win32_pnpentity中不會列出?藏的設備, 無法用win32_pnpentity來作.

param($runadmin)


function remove-HiddenDevice {

  $Username = "vhcy\vhwcmis"
  $Password = "Mis20190610"
  $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
  $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

  $dev =  Get-PnpDevice | Where-Object -FilterScript {$_.Present -eq $false -and $_.Class -in ('SmartCard','SmartCardReader','SmartCardFilter')}

  # 部分舊win10系統旳pnputil.exe沒有remove-device, 所以copy了一份, 再跟os裡的比一下新舊.
  $pnputil = "$PSCommandPath\pnputil.exe"
  $pnputil_os = "C:\Windows\system32\pnputil.exe"
  $result = (Get-ItemPropertyValue -Path $pnputil -name VersionInfo).productVersion -lt (Get-ItemPropertyValue -Path $pnputil_os -name VersionInfo).productVersion
  if ($result) {$pnputil = $pnputil_os}

  # todo:  win7要用devcon.exe, win10也有但要另安裝windows10 WDK
  # Win7的要再測看看.


  foreach ($d in $dev) {
    Start-Process -FilePath $pnputil -ArgumentList "/remove-device $($d.instanceID)" -Credential $credential -Wait -NoNewWindow
  }

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