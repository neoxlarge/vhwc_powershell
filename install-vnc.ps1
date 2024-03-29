# 安裝VNC
# setup.exe /? 可以顯示安裝指令

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-VNC {
    
    $software_name = "UltraVnc"
    $software_path = "\\172.20.5.187\mis\08-VNC\1_2_24"
    $software_msi = "UltraVNC_1_2_24_X64_Setup.exe"
    $software_msi_x86 = "UltraVNC_1_2_24_X86_Setup.exe"
    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到本機暫存"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi }
            "x86" { $software_exec = $software_msi_x86 }
            default { Write-Warning "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($null -ne $software_exec) {
            Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/silent /loadinf=$($env:temp + "\" + $software_path.Name + "\installvnc.inf")" -Wait
            Start-Sleep -Seconds 5 
        }
        else {
            Write-Warning "$software_name 無法正常安裝."
        }
      
        #復制設定檔vltravnc.ini 到C:\Program Files\uvnc bvba\UltraVNC
        Copy-Item -Path ($env:temp + "\" + $software_path.Name + "\ultravnc.ini") -Destination ($env:ProgramFiles + "\uvnc bvba\UltraVNC") -Force

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
        install-VNC    
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}


