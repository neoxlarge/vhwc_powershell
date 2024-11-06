# 安裝CMS_CGServiSignAdapter
# 20240611 因為CGServiSignAdapter有更新到1.0.23
# 原本的1.0.22會要求要更新才能登入健保VPN.

param($runadmin)

# 要求powershell v5.1以上才執行, win7預設powershell v2.0.
if (!$PSVersionTable.PSCompatibleVersions -match "^5\.1") {
    Write-Output "powershell requires version 5.1, exit"
    Start-Sleep -Seconds 3
    exit
}

function import-vhwcmis_module {
    $moudle_paths = @(
        if ($script:MyInvocation.MyCommand.Path) {"$(Split-Path $script:MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue)"},
        "d:\mis\vhwcmis",
        "c:\mis\vhwcmis",
        "\\172.20.5.185\powershell\vhwc_powershell",
        "\\172.20.1.14\share\00新電腦安裝\vhwc_powershell",
        "\\172.20.1.122\share\software\00newpc\vhwc_powershell",
        "\\172.19.1.229\cch-share\h040_張義明\vhwc_powershell"
    )

    $filename = "vhwcmis_module.psm1"

    foreach ($path in $moudle_paths) {
        
        if (Test-Path "$path\$filename") {
            write-output "$path\$filename"
            Import-Module "$path\$filename" -ErrorVariable $err_import_module
            if ($err_import_module -eq $null) {
                Write-Output "Imported module path successed: $path\$filename"
                break
            }
        }
    }

    $result = get-command -Name "get-installedprogramlist" -CommandType Function -ErrorAction SilentlyContinue
    if ($result -eq $null) {
        throw "無法載入vhwcmis_module模組, 程式無法正常執行. "
    }
}
import-vhwcmis_module


function install-CMS {
    $log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VHWC_logs\update-cms.log"

    ## 安裝CMS_CGServiSignAdapter
    ### 依文件要求,安裝前應關閉防毒軟體, 所以比防毒先安裝

    $software_name = "NHIServiSignAdapterSetup"
    $software_path = "\\172.20.5.187\mis\05-CMS_CGServiSignAdapterSetup\CMS_CGServiSignAdapterSetup"
    $software_exec = "NHIServiSignAdapterSetup.exe"
    
    $all_installed_program = get-installedprogramlist
   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }


    if ($software_is_installed) {
        $exe_version = (Get-ItemProperty -Path "$software_path\$software_exec").VersionInfo.FileversionRaw.toString()

        $result = Compare-Version -Version1 $exe_version -Version2 $software_is_installed.DisplayVersion

        if ($result) {
 
            $ipv4 = Get-IPv4Address 
            Write-Log -logfile $log_file -message "Find old CMS version:$($software_is_installed.DisplayVersion)"
  
            Write-Output "Find old CMS $software_name, version: $($software_is_installed.DisplayVersion)"
            Write-Output "Removing old version."
            Start-Process -FilePath $software_is_installed.UninstallString -ArgumentList "/S" -Wait
            $software_is_installed = $null
        }
    }

    if ($software_is_installed -eq $null ) {
        # 沒安裝, 直接安裝.
        Write-Output "Start to install $software_name"

        #來源路徑 ,要復制的路徑,and 安裝執行程式名稱
        $software_path = get-item -Path $software_path
                
        #復制檔案到temp
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        #installing...
        $process_id = Start-Process -FilePath "$env:temp\$($software_path.Name)\$software_exec" -PassThru

        #依安裝文件, CGServiSignMonitor會最後被開啟, 所以檢查到該程序執行後, 表示安裝完成.
        $process_exist = $null
        while ($process_exist -eq $null) {
            $process_exist = Get-Process -Name CGServiSignMonitor -ErrorAction SilentlyContinue
            if ($process_exist -ne $null) { Stop-Process -Name $process_id.Name }
            write-output ($process_id.Name + "is installing...wait 5 seconds.")
            Start-Sleep -Seconds 5
        } 

        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }


    }

    Write-output ("software has installed:" + $software_is_installed.DisplayName )
    Write-Output ("Version:" + $software_is_installed.DisplayVersion)


    

}



#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }
    if ($check_admin) {
    install-CMS
    } else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試." 
    }
    pause
    #Start-Sleep -Seconds 5
}