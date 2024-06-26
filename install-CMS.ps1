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


function Get-IPv4Address {
    <#
    回傳找到的IP,只能在172.*才能用. 
    #>

    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.20.*" } |
    Select-Object -ExpandProperty IPAddress |
    Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
    Select-Object -First 1

    if ($ip -eq $null) {
        return $null
    }
    else {     
        return $ip
    }
}

function install-CMS {

    # 取得vhwcmis_module.psm1的3種方式:
    # 1.程式執行當前路徑, 放到Group police執行可能抓不到.
    # 2.常用的路徑, d:\mis\vhwc_powershell, 不是每台都有放.
    # 3.連到NAS上取得. 非網域的電腦會沒有NAS的權限, 須手動連上NAS.

    $pspaths = @()
    if ($PSCommandPath -contains "\") {$pspaths += "$(Split-Path $PSCommandPath)\vhwcmis_module.psm1"}

    $local_path = "d:\mis\vhwc_powershell\vhwcmis_module.psm1"
    if (Test-Path $local_path){$pspaths += $local_path}

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

    foreach ($path in $pspaths) {
        Import-Module $path -ErrorAction SilentlyContinue
        if ((get-command -Name "get-installedprogramlist" -CommandType Function -ErrorAction SilentlyContinue)) {
            break
        }
    }


    $log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\install-cms.log"
    

    ## 安裝CMS_CGServiSignAdapter
    ### 依文件要求,安裝前應關閉防毒軟體, 所以比防毒先安裝

    $software_name = "NHIServiSignAdapterSetup"
    $software_path = "\\172.20.1.122\share\software\00newpc\05-CMS_CGServiSignAdapterSetup\CMS_CGServiSignAdapterSetup"
    $software_exec = "NHIServiSignAdapterSetup.exe"
    
    $all_installed_program = get-installedprogramlist
   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }


    if ($software_is_installed) {
        $exe_version = (Get-ItemProperty -Path "$software_path\$software_exec").VersionInfo.FileversionRaw.toString()

        $result = Compare-Version -Version1 $exe_version -Version2 $software_is_installed.DisplayVersion

        if ($result) {
            $ipv4 = Get-IPv4Address 

            $log_string = "Find old CMS version:$($software_is_installed.DisplayVersion),$env:COMPUTERNAME,$ipv4,$(Get-Date)"
            $log_string | Add-Content -PassThru $log_file
            
  
            Write-Output "Find old CMS version $software_name : $($software_is_installed.DisplayVersion)"
            Write-Output "Removing old version."
            Start-Process -FilePath $software_is_installed.UninstallString -ArgumentList "/S" -Wait
            $software_is_installed = $null
        }
    }

    if ($software_is_installed -eq $null) {
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
  
    install-CMS
  
    #pause
    Start-Sleep -Seconds 5
}