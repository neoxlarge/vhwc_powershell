# 安裝工具程式7z
# 20242024 己改成會自動移除舊的版本.

param($runadmin)

function import-vhwcmis_module {
    $moudle_paths = @(
        "$(Split-Path $script:MyInvocation.MyCommand.Path)",
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
                Write-Output "Imported module path successed: $path"
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

function install-7z {
    
    if (!$credential) {
        $credential = get-admin_cred
    }

    $software_name = "7-Zip*"
    $software_path = "\\172.20.5.187\mis\07-7-7z"
    $software_msi = "7z-x64.msi"
    $software_msi_x86 = "7z-x32.msi"

    ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $software_exec = $software_msi }
        "x86" { $software_exec = $software_msi_x86 }
        default { throw "$software_name 無法正常安裝: 不支援的系統:  $($env:PROCESSOR_ARCHITECTURE)" }
    }

    $msi_version = get-msiversion -MSIPATH $($software_path + "\" + $software_exec)
    Write-output "========================"
    Write-Output "Software: $software_name"
    Write-Output "Source path: $software_path\$software_exec"
    write-output "Source version: $msi_version"
    Write-output "========================"


    # 預先連線到安裝來源的路徑.
    if (!(Test-Path -Path $software_path)) {
        New-PSDrive -Name $software_name -Root "$software_path" -PSProvider FileSystem -Credential $credential
        }

    ## 找出軟體是否己安裝
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed) {
        #己有安裝
        # 比較版本新舊
        $msi_version = get-msiversion -MSIPATH $($software_path + "\" + $software_exec)
        $check_version = compare-version -Version1 $msi_version -Version2 $software_is_installed.DisplayVersion

        if ($check_version) {
            #msi版本比較新,移除舊的後, 重新查詢$software_is_installed
            Write-Output "Find installed version(old): $($software_is_installed.DisplayVersion)"
            Write-Output "Uninstall string: $($software_is_installed.uninstallString)"
            Write-Output "Quick Uninstalling string : $($software_is_installed.QuietUninstallString)"
            Write-Output "Uninstalling $software_name"
            Write-output "========================"
            uninstall-software -name $software_is_installed.DisplayName

            
            $all_installed_program = get-installedprogramlist
            $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
        }

    } 
    

    if ($null -eq $software_is_installed) {
    
        Write-Output "Insalling: $software_name"

        #復制檔案到本機暫存"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        
        if ($software_exec -ne $null) {
            $msiExecArgs = "/i $($env:temp + "\" + $software_path.Name + "\" + $software_exec) /passive"
            
            if ($check_admin) {
                # 有管理員權限
                $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiExecArgs -PassThru
            }
            else {
                # 無管理員權限
                $credential = get-admin_cred
                $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiExecArgs -Credential $credential -PassThru
            }
            
            $proc.WaitForExit()
        } 
      
        
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

    install-7z    

    pause
}