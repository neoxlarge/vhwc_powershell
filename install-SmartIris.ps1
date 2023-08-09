# 安裝SmartIris

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-SmartIris {
    
    $software_name = "SmartIris"
    $software_path = "\\172.20.5.187\mis\02-SmartIris\SmartIris_V1.3.6.4_Beta7_UQ-1.1.0.19_R2_Install_20200701"
    $software_exe = "setup.exe"
    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到本機暫存"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        # 安裝  
        $runid = Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exe) -ArgumentList "/s /f1$($env:temp + "\" + $software_path.Name + "\vhwc.iss")" -PassThru
        
        # 安裝過程中, 這2支程會跳出來, 先砍了, 等會再設定即可.
        while (!($runid.HasExited)) {
            get-process -Name MonitorCfg -ErrorAction SilentlyContinue | Stop-Process
            get-process -Name UQ_Setting -ErrorAction SilentlyContinue | Stop-Process
            Start-Sleep -Seconds 1
        }

        #依照 \\172.19.1.14\Update\資訊室\共用程式\SmartIris\SmartIris更版SOP.doc
        #步驟三：
        #\\172.19.1.14\Update\資訊室\共用程式\SmartIris\UltraQuery_V1.1.1.0_Update_20200731
        # 全選複製到C:\TEDPC\SmartIris\UltraQuery並覆蓋

        copy-item -Path "\\172.20.5.187\mis\02-SmartIris\UltraQuery_V1.1.1.0_Update_20200731\*" -Destination "C:\TEDPC\SmartIris\UltraQuery" -Recurse -Force

        #復制設定檔到本機.
        copy-item -Path "\\172.20.5.187\mis\02-SmartIris\vhwc_UltraQuery_SysIni\*" -Destination "C:\TEDPC\SmartIris\UltraQuery\SysIni" -Force
 

        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    
    }

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)

}



#檔案獨立執行時會執行函式, 如果是被載入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        install-SmartIris    
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}