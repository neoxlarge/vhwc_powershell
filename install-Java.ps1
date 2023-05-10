# 安裝Java runtime 

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function set-Java_env {
    Write-Output "取消JAVA自動檢查更新"
    
    #寫入的值是binary型態
    $binary_value = @(0x01, 0x00, 0x00, 0x00, 0xD0, 0x8C, 0x9D, 0xDF, 0x01, 0x15, 0xD1, 0x11,
        0x8C, 0x7A, 0x00, 0xC0, 0x4F, 0xC2, 0x97, 0xEB, 0x01, 0x00, 0x00, 0x00,
        0xC4, 0x0D, 0xE3, 0xD1, 0x21, 0xEC, 0x3B, 0x4B, 0x99, 0x6E, 0xCD, 0x4B,
        0x90, 0xBB, 0xBD, 0x87, 0x00, 0x00, 0x00, 0x00, 0x1C, 0x00, 0x00, 0x00,
        0x50, 0x00, 0x61, 0x00, 0x73, 0x00, 0x73, 0x00, 0x77, 0x00, 0x6F, 0x00,
        0x72, 0x00, 0x64, 0x00, 0x20, 0x00, 0x44, 0x00, 0x61, 0x00, 0x74, 0x00,
        0x61, 0x00, 0x00, 0x00, 0x03, 0x66, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00,
        0x10, 0x00, 0x00, 0x00, 0xA3, 0xA8, 0x42, 0xB2, 0xE9, 0xB3, 0xDE, 0xC6,
        0x27, 0x52, 0x6E, 0x40, 0x71, 0xCB, 0x34, 0x53, 0x00, 0x00, 0x00, 0x00,
        0x04, 0x80, 0x00, 0x00, 0xA0, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00,
        0x50, 0x0F, 0x30, 0xFB, 0xEC, 0x55, 0x7D, 0x4B, 0x3E, 0x66, 0xF8, 0xF1,
        0xF9, 0xCA, 0x76, 0xA9, 0x08, 0x00, 0x00, 0x00, 0x3A, 0x0F, 0xEC, 0x0C,
        0xAC, 0xB5, 0x35, 0x4B, 0x14, 0x00, 0x00, 0x00, 0x26, 0x4B, 0x3E, 0xA9,
        0x9F, 0x3F, 0xE2, 0x35, 0x20, 0xF8, 0x53, 0xF2, 0xF4, 0x47, 0x14, 0xF3,
        0xA1, 0xF7, 0x9F, 0x27)
    New-Item -Path "HKCU:\SOFTWARE\JavaSoft\Java Update\Policy" -ItemType "Directory" -Force
    New-ItemProperty -Path "HKCU:\SOFTWARE\JavaSoft\Java Update\Policy" -Name "EnableAutoUpdateCheck" -PropertyType "Binary" -Value $binary_value -Force
    
    Write-Output "修改JAVA執行參數,加入-Xmx256m"
    <# 
    要自動修改 Java 控制台設定，可以透過 PowerShell 修改 Java 控制台配置文件 deployment.properties。
    該文件包含了 Java 控制台的各種配置選項，您可以透過修改該文件來自動化更改 Java 控制台的設定。
    #>
    $filePath = "$env:USERPROFILE\AppData\LocalLow\Sun\Java\Deployment\deployment.properties"

    # 檢查文件是否存在
    if (Test-Path $filePath) {
        # 讀取文件內容
        $content = Get-Content $filePath -Raw
    
        # 將 deployment.javaws.jre.0.args 選項設置為 -Xmx256m
        $content = $content -replace "deployment.javaws.jre.0.args=.*", "deployment.javaws.jre.0.args=-Xmx256m"
    
        # 將修改後的內容寫回文件
        Set-Content $filePath $content -Force
    }
    else {
        Write-Error "找不到 deployment properties 文件。"
    }
}

function install-Java {

    $software_name = "Java*"
    $software_path = "\\172.20.5.187\mis\11-中榮系統\01-中榮系統環境設定\中榮系統環境設定\java"
    $software_msi = "jre-6u13-windows-i586-p.exe"

    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到本機暫存"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force

       
        Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_msi) -ArgumentList "/passive" -Wait
        Start-Sleep -Seconds 5 
                    
        
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
        install-Java    
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    
    set-Java_env
    pause
}