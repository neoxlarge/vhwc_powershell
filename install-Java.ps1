# 安裝Java runtime 

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function set-Java_env {
    Write-Output "取消JAVA自動檢查更新"
    
    #參考連結 https://thegeekpage.com/turn-off-java-update-notification-in-windows-10/
    #32bit/64bit系統要修改的registry位置不一樣.
    #修改正確會把更新的分頁完全關閉不顯示.
    #沒有admin權限不用執行取消更新動作.

    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $reg_path = "HKLM:\SOFTWARE\WOW6432Node\JavaSoft\Java Update\Policy" }
        "x86" { $reg_path = "HKLM:\SOFTWARE\JavaSoft\Java Update\Policy" }
    }
    
    
    if ((Get-ItemProperty -Path $reg_path -Name "EnableJavaUpdate").enableJavaupdate -ne 0 ) 
    {
        if ($check_admin) {
            Set-ItemProperty -Path $reg_path -Name "EnableJavaUpdate" -Value 0 -Force
        } else {
            Write-Warning "Java己啟用自動更新, 但沒有管理者權限修改, 請用管理者重試."
        }
    }  else {
        Write-Output "Java更新己取消"
    }

    <# 取消整個更新分頁後,不再需要以下設定.
    Write-Output "修改Java更新通知為下載之前"
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" {
            Set-ItemProperty -Path "HKLM:\Software\WOW6432Node\JavaSoft\Java Update\Policy" -name "NotifyInstall" -Value 0 -Force
            Set-ItemProperty -Path "HKLM:\Software\WOW6432Node\JavaSoft\Java Update\Policy" -name "NotifyDownload" -Value 1 -Force
        }
    
        "x86" {
            Set-ItemProperty -Path "HKLM:\Software\JavaSoft\Java Update\Policy" -name "NotifyInstall" -Value 0 -Force
            Set-ItemProperty -Path "HKLM:\Software\JavaSoft\Java Update\Policy" -name "NotifyDownload" -Value 1 -Force
        }
    }
    #>


    Write-Output "修改JAVA執行參數,加入-Xmx256m"
    <# 
    要自動修改 Java 控制台設定，可以透過 PowerShell 修改 Java 控制台配置文件 deployment.properties。
    該文件包含了 Java 控制台的各種配置選項，您可以透過修改該文件來自動化更改 Java 控制台的設定。
    #>
    $filePath = "$env:USERPROFILE\AppData\LocalLow\Sun\Java\Deployment\deployment.properties"

    #在剛安裝完JAVA時, 未執行JAVA控制台時, deployment.properties檔案會不在, 執行JAVA控制台會產生該檔.
    if (!(test-path $filePath)) {
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" {Start-Process -FilePath "C:\Program Files (x86)\Java\jre6\bin\javacpl.exe"}
            "x86" {Start-Process -FilePath "C:\Program Files\Java\jre6\bin\javacpl.exe"}
        }
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