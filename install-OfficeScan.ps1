# 安裝防毒 Trend Micro Apex One Security Agent

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


function install-OfficeScan {
    # 安裝防毒 Trend Micro Apex One Security Agent
    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist
  
    $software_name = "Trend Micro Apex One Security Agent"
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到"D:\mis"
        $software_path = get-item -Path "\\172.20.1.14\share\software\officescan_antivir"
        if (Test-Path -Path "d:\mis") {
            $software_copyto_path = "D:\mis"
        }
        else {
            $software_copyto_path = "C:\mis"
        }
        
    
        Copy-Item -Path $software_path -Destination $software_copyto_path -Recurse -Force

        ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = "agent_cloud_x64.msi" }
            "x86" { $software_exec = "agent_cloud_x86.msi" }
            default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($software_exec -ne $null) {
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $($software_copyto_path + "\" + $software_path.Name + "\" + $software_exec) /passive /log d:\mis\install_officescan.log" -Wait
            Start-Sleep -Seconds 5 
        }
        else {
            $warn_msg += "Software install fail: $software_name"
            Write-Warning $warn_msg[-1] 
        }
      
        #安裝完, 刪除安裝檔案
        Remove-Item -Path ($software_copyto_path + "\" + $software_path.Name) -Recurse -Force
      
        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    
    }

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)

}



#檔案獨立執行時會執行函式, 如果是被?入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        install-officescan
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}