# 安裝工具程式7z
# 20242024 己改成會自動移除舊的版本.

param($runadmin)

Import-Module -name "$(Split-Path $PSCommandPath)\vhwcmis_module.psm1"

function install-7z {
    
    $software_name = "7-Zip*"
    $software_path = "\\172.20.5.187\mis\07-7-7z"
    $software_msi = "7z2201-x64.msi"
    $software_msi_x86 = "7z2201-x32.msi"

    ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $software_exec = $software_msi }
        "x86" { $software_exec = $software_msi_x86 }
        default { throw "$software_name 無法正常安裝: 不支援的系統:  $($env:PROCESSOR_ARCHITECTURE)" }
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
            Write-Output "找到舊的版本: $($software_is_installed.DisplayName) : $($software_is_installed.DisplayVersion)"
            Write-Output "Uninstall string: $($software_is_installed.uninstallString)"
            uninstall-software -name $software_is_installed.DisplayName
            
            $all_installed_program = get-installedprogramlist
            $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
        }

    } 
    

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

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