
param($runadmin)


function disable-sleep {
    Write-Output "變更電源計畫:"

    write-output "設定電源計劃:平衡"
    powercfg /setactive "SCHEME_BALANCED"

    Write-Output "硬碟-關閉硬碟前的時間:0"
    powercfg /change disk-timeout-ac 0

    write-output  "關閉顯示器:15分"
    powercfg /change monitor-timeout-ac 15
    
    write-output "讓電腦睡眠:永不"
    powercfg /change standby-timeout-ac 0

    Write-Output "關閉混合式睡眠:關閉"
    powercfg /hibernate off

}

#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    disable-sleep
    pause
}