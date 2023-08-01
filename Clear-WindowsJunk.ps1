#清理Windows的暫存和回收桶
param($runadmin)


function Clear-WindowsJunk {
  # 清除用戶暫存檔 
  Write-Output "清除用戶暫存檔$env:LOCALAPPDATA\Temp\*"
  Remove-Item -Path "$env:LOCALAPPDATA\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue

  Write-Output "清除系統暫存檔env:windir\Temp\*"
  # 清除系統暫存檔 
  Remove-Item -Path "$env:windir\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue

  # 清空回收桶 
  Write-OutPut "清空回收桶"
  Clear-RecycleBin -Force -ErrorAction SilentlyContinue

  # 清除事件查看器日誌 
  #WEVTUtil.exe cl Application
  #WEVTUtil.exe cl System
  #WEVTUtil.exe cl Security
}

  
#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Clear-WindowsJunk
    pause
}