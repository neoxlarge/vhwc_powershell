#install IE11

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-IE11 {

    $software_name = "IE11"
    $software_path = "\\172.20.5.187\mis\20-IE"
    $software_msi_x64 = "x64\IE11-Windows6.1-x64-zh-tw.exe"
    $software_msi_x32 = "x86\IE11-Windows6.1-x86-zh-tw.exe"


    # 取得目前系統中 IE 的版本
    #此命令將嘗試從註冊表的 "HKLM:\Software\Microsoft\Internet Explorer" 位置中檢索 IE 版本資訊，並將其存儲在 $ieVersion 變數中。
    #請注意，使用 svcVersion 或 Version 屬性取決於 IE 的版本。svcVersion 屬性適用於 IE 10 或更新版本，而 Version 屬性則適用於 IE 9 或較舊版本。
    $ieVersion = $null
    $ieVersion = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Internet Explorer" | Select-Object -Property "svcVersion").svcVersion
    if ($ieVersion -eq $null) {
        $ieVersion = (Get-ItemPropertyValue -Path "HKLM:\Software\Microsoft\Internet Explorer" -Name "Version")
    }

    Write-Output "IE version: $ieVersion"
    
    # 檢查版本並安裝 IE 11（如果版本小於 11）
    if ([int16]$ieVersion.Split(".")[0] -lt 11) {
        
        Write-Output "Start to insall: $software_name"

        #復制檔案到本機暫存"
        $software_path = get-item -Path $software_path
        if (!(Test-Path "$env:temp\$($software_path.Name)")) {
            Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose
        }

        ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-Warning "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($null -ne $software_exec) {
            Start-Process -FilePath "$env:temp\$($software_path.Name)\$software_exec" -ArgumentList "/passive /norestart" -Wait
            Start-Sleep -Seconds 5 

            Write-Output "IE 11 安裝結束."
            write-output "安裝IE 11的hotfix"

            # 安裝 MSU（Microsoft Update Standalone Package）安裝包，
            # 您可以使用 Start-Process 命令來執行 wusa.exe（Windows Update Standalone Installer）工具，並指定要安裝的 MSU 檔案。
            
            $hotfix = get-childitem -Path "$env:temp\$($software_path.Name)\$($software_exec.Split("\")[1])\*" -Include "*.msu"

            foreach ($h in $hotfix) {
                Write-Output "Installing hotfix: $($h.Name) "
                Start-Process -FilePath wusa.exe -ArgumentList "$($h.fullname) /quiet /norestart" -Wait
                Start-Sleep -Seconds 3 
            }

        }
        else {
            Write-Warning "$software_name 無法正常安裝."
        }
        
    }

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
        install-IE11
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}