# 安裝SSLVPN用軟體
# SMA connect Ageng 和 Net Extender

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


function install-SMAConnectAgent {


    $software_name = "SMA Connect Agent"
    $software_path = "\\172.20.1.122\share\software\00newpc\34-SSLVPN"
    $software_msi_x64 = "SMAConnectAgent.msi"  #64bit 32bit 都同一個
    $software_msi_x32 = "SMAConnectAgent.msi"

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $all_installed_program = get-installedprogramlist
      
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到temp
        $software_path = get-item -Path $software_path
                
        #copy-item 無法接認證, 須要從psdrive接, 所以要掛driver.
        $net_driver = "vhwcdrive" #只是給個driver名字而己.
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        Copy-Item -Path "$($net_driver):\" -Destination $env:TEMP -Recurse -Force -Verbose 
        Remove-PSDrive -Name $net_driver


        ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($software_exec -ne $null) {
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $($env:temp + "\" + $software_path.Name + "\" + $software_exec) /passive /log install_officescan.log" -Wait
            Start-Sleep -Seconds 5 
        }
        else {
            Write-Warning "$software_name 無法正常安裝."
        }
      
             
        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    
    }

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)

}

function install-NX {
    #

    $software_name = "SonicWall NetExtender"
    $software_path = "\\172.20.1.122\share\software\00newpc\34-SSLVPN"
    $software_msi_x64 = "NXSetupU.exe"  #64bit 32bit 都同一個
    $software_msi_x32 = "NXSetupU.exe"

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    $all_installed_program = get-installedprogramlist
      
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到temp
        $software_path = get-item -Path $software_path
                
        #copy-item 無法接認證, 須要從psdrive接, 所以要掛driver.
        $net_driver = "vhwcdrive" #只是給個driver名字而己.
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        Copy-Item -Path "$($net_driver):\" -Destination $env:TEMP -Recurse -Force -Verbose 
        Remove-PSDrive -Name $net_driver


        ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi_x64 }
            "x86" { $software_exec = $software_msi_x32 }
            default { Write-OutPut "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

        if ($software_exec -ne $null) {
            Start-Process -FilePath "$($env:temp)\$($software_path.name)\$software_exec"    -ArgumentList "/S" -Wait
            Start-Sleep -Seconds 5 
        }
        else {
            Write-Warning "$software_name 無法正常安裝."
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

    if ($check_admin) { 
        install-SMAConnectAgent
        install-NX
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}