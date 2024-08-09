## 安裝 HCAServiSign

param($runadmin)

# 要求powershell v5.1以上才執行, win7預設powershell v2.0.
if (!$PSVersionTable.PSCompatibleVersions -match "^5\.1") {
    Write-Output "powershell requires version 5.1, exit"
    Start-Sleep -Seconds 3
    exit
}

function import-vhwcmis_module {
    # import vhwcmis_module.psm1
    # 取得vhwcmis_module.psm1的3種方式:
    # 1.程式執行當前路徑, 放到AD上用Group police執行,不會有當前路徑.
    # 2.常用的路徑, d:\mis\vhwc_powershell, 不是每台都有放.
    # 3.連到NAS上取得. 非網域的電腦會沒有NAS的權限, 須手動連上NAS.

    $pspaths = @()

    if ($script:MyInvocation.MyCommand.Path -ne $null) {
        $work_path = "$(Split-Path $script:MyInvocation.MyCommand.Path)\vhwcmis_module.psm1"
        if (test-path -Path $work_path) { $pspaths += $work_path }
    }
    $nas_name = "nas122"
    $nas_path = "\\172.20.1.122\share\software\00newpc\vhwc_powershell"
    if (!(test-path $nas_path)) {
        $nas_Username = "software_download"
        $nas_Password = "Us2791072"
        $nas_securePassword = ConvertTo-SecureString $nas_Password -AsPlainText -Force
        $nas_credential = New-Object System.Management.Automation.PSCredential($nas_Username, $nas_securePassword)
        
        New-PSDrive -Name $nas_name -Root "$nas_path" -PSProvider FileSystem -Credential $nas_credential | Out-Null
    }
    $pspaths += "$nas_path\vhwcmis_module.psm1"

    $local_path = "d:\mis\vhwc_powershell\vhwcmis_module.psm1"
    if (Test-Path $local_path) { $pspaths += $local_path }

    foreach ($path in $pspaths) {
        Import-Module $path -ErrorAction SilentlyContinue
        if ((get-command -Name "get-installedprogramlist" -CommandType Function -ErrorAction SilentlyContinue)) {
            break
        }
    }
}
import-vhwcmis_module


function install-HCA {
    
    $log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VHWC_logs\update-HCA.log"

    ## 安裝 HCAServiSign
    ### 找出軟體是否己安裝

    $software_name = "HCAServiSignAdapterSetup"
    $software_path = "\\172.20.1.122\share\software\00newpc\05-HCAServiSign醫事卡解鎖"
    $software_exec = "HCAServiSignAdapterSetup.exe"

    $all_installed_program = get-installedprogramlist
  
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    if ($software_is_installed) {
        $exe_version = (Get-ItemProperty -Path "$software_path\$software_exec").VersionInfo.FileversionRaw.toString()
        $result = Compare-Version -Version1 $exe_version -Version2 $software_is_installed.DisplayVersion

        if ($result) {
            $ipv4 = Get-IPv4Address 
            Write-Log -logfile $log_file -message "Find old HCA version:$($software_is_installed.DisplayVersion)"
  
            Write-Output "Find old HCA $software_name, version: $($software_is_installed.DisplayVersion)"
            Write-Output "Removing old version."
            Start-Process -FilePath $software_is_installed.UninstallString -ArgumentList "/S" -Wait
            $software_is_installed = $null

        }
    }

    if ($null -eq $software_is_installed) {
        Write-Output "Start to install $software_name"

        #來源路徑 ,要復制的路徑,and 安裝執行程式名稱
        $software_path = get-item -Path $software_path
        
        #復制檔案到暫存資料夾
        Copy-Item -Path $software_path -Destination $env:TEMP -Recurse -Force -Verbose

        #installing...
        $process_id = Start-Process -FilePath ($env:TEMP + "\" + $software_path.Name + "\" + $software_exec) -PassThru
    
        #依安裝文件, HCAServiSignMonitor會最後被開啟, 所以檢查到該程序執行後, 表示安裝完成.
        $process_exist = $null
        while ($process_exist -eq $null) {
            $process_exist = Get-Process -Name HCAServiSignMonitor -ErrorAction SilentlyContinue
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

    if (! ($(Get-PSDrive -Name $nas_name -ErrorAction SilentlyContinue) -eq $null) ) {
        Remove-PSDrive -Name $nas_name
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

    if ($check_admin) { 
        install-HCA
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    #pause
    Start-Sleep -Seconds 5
}
