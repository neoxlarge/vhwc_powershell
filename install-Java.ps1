# 安裝Java runtime 

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function set-Java_env {
    Write-Output "取消JAVA自動檢查更新"
    
    #寫入的值是binary型態
    $binary_value = @(0x01,0x00,0x00,0x00,0xD0,0x8C,0x9D,0xDF,0x01,0x15,0xD1,0x11,0x8C,0x7A,0x00,0xC0,0x4F,0xC2,0x97,0xEB,0x01,0x00,0x00,0x00,0xDF,
    0xBA,0x98,0x52,0x33,0x8F,0x96,0x48,0xA0,0x1C,0x68,0x71,0xE4,0x07,0x99,0xF0,0x00,0x00,0x00,0x00,0x1C,0x00,0x00,0x00,0x50,0x00,
    0x61,0x00,0x73,0x00,0x73,0x00,0x77,0x00,0x6F,0x00,0x72,0x00,0x64,0x00,0x20,0x00,0x44,0x00,0x61,0x00,0x74,0x00,0x61,0x00,0x00,
    0x00,0x10,0x66,0x00,0x00,0x00,0x01,0x00,0x00,0x20,0x00,0x00,0x00,0xB5,0x39,0x2B,0xE7,0xD3,0x9D,0xCE,0xEC,0x98,0x7B,0x14,0x9B,
    0xA1,0xC0,0xD7,0xAE,0xBD,0xF3,0xFC,0xF2,0xE3,0xE4,0x19,0x09,0x5B,0xFC,0x2F,0xB5,0xAA,0x71,0x26,0x88,0x00,0x00,0x00,0x00,0x0E,
    0x80,0x00,0x00,0x00,0x02,0x00,0x00,0x20,0x00,0x00,0x00,0x43,0x04,0x60,0x03,0xAB,0x05,0x2E,0xC2,0x1D,0xC9,0xCB,0x74,0xF7,0xA0,
    0x80,0x20,0xC2,0x91,0xEC,0x0E,0x2A,0x2E,0x5D,0x66,0x57,0xCA,0xFE,0x69,0x8D,0x65,0x44,0xFE,0x10,0x00,0x00,0x00,0x0F,0x3D,0x63,
    0xFF,0x83,0xEA,0x92,0x63,0xBA,0x82,0x67,0xAE,0x99,0x02,0xD0,0xF0,0x40,0x00,0x00,0x00,0x92,0x42,0x6E,0x54,0xFD,0x08,0x40,0xAB,
    0x35,0x61,0x4A,0x24,0x28,0x47,0xD2,0xC5,0x2A,0x76,0x61,0x90,0xF3,0xED,0x68,0x5A,0x33,0x2A,0x54,0xB2,0xCD,0x95,0xF1,0x27,0x37,
    0x7C,0x69,0x2A,0xAF,0x96,0x77,0x95,0x54,0x83,0xBF,0xC9,0x85,0xBB,0x81,0x35,0x0D,0x6B,0xBA,0x76,0xB6,0xF3,0xE4,0x68,0xCD,0xB1,
    0x3E,0xBE,0x41,0xE5,0xE1,0xAB)
    New-Item -Path "HKCU:\SOFTWARE\JavaSoft\Java Update\Policy" -ItemType "Directory" -Force
    New-ItemProperty -Path "HKCU:\SOFTWARE\JavaSoft\Java Update\Policy" -Name "EnableAutoUpdateCheck" -PropertyType "Binary" -Value $binary_value -Force
    
    Write-Output "修改通知為下載之前"

    Set-ItemProperty -Path "HKLM:\Software\WOW6432Node\JavaSoft\Java Update\Policy" -name "NotifyInstall" -Value 0 -Force
    Set-ItemProperty -Path "HKLM:\Software\WOW6432Node\JavaSoft\Java Update\Policy" -name "NotifyDownload" -Value 1 -Force


    Write-Output "修改JAVA執行參數,加入-Xmx256m"
    <# 
    要自動修改 Java 控制台設定，可以透過 PowerShell 修改 Java 控制台配置文件 deployment.properties。
    該文件包含了 Java 控制台的各種配置選項，您可以透過修改該文件來自動化更改 Java 控制台的設定。
    #>
    $filePath = "$env:USERPROFILE\AppData\LocalLow\Sun\Java\Deployment\deployment.properties"

    #在剛安裝完JAVA時, 未執行JAVA控制台時, deployment.properties檔案會不在, 執行JAVA控制台會產生該檔.
    if (!(test-path $filePath)) {
        Start-Process -FilePath "C:\Program Files (x86)\Java\jre6\bin\javacpl.exe"
        Start-Sleep -Seconds 5
        Stop-Process -Name "javaw"
    }

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