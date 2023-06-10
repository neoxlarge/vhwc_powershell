
#啟用powershell遠端管理
<#
PowerShell 遠端管理允許 IT 專業人員從遠程地點管理多台計算機，節省了時間和資源。
使用 PowerShell 遠端管理，可以輕鬆地從本地計算機遠程執行 PowerShell 腳本，
查看計算機狀態，設置系統配置，安裝軟體和更新，檢查事件日誌，還可以執行其他管理任務。
這使得 IT 管理員能夠更快地處理問題，減少了出差和現場工作的需求，並提高了效率。此外，
PowerShell 遠端管理支援跨平台操作，可以在 Windows、Linux 和 macOS 上使用。

遠端管理(WinRM)須加入AD才能enable, 非AD環境須額外設定,較為麻煩.
#>
param($runadmin)

Function Check-EnablePSRemoting {
    $winrmService = Get-Service -Name WinRM
    $isJoinAD = (Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain
    If ($winrmService.Status -ne "Running") {

        if ($check_admin -and $isJoinAD) {
            Write-Output "PowerShell 遠端管理未啟用，現在正在啟用..."
            Enable-PSRemoting
            Write-Output "PowerShell 遠端管理已啟用."
        } else {
            Write-Warning "沒有系統管理員權限或未加入AD,且 PowerShell 遠端管理未啟用,請以系統管理員身分重新嘗試."
        }
    }
    Else {
        Write-Output "PowerShell 遠端管理已經啟用."
    }
}


 
#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
    #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
    Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Check-EnablePSRemoting
    pause
}