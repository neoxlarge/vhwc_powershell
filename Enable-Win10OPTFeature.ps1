﻿#Enable-SMBv1
<#
要連線到172.20.1.14, 需要啟用SMBv1/CIFS功能.

SMB（Server Message Block）和CIFS（Common Internet File System）
是用於共享文件和印表機的網絡協議。SMB最初由IBM和Microsoft開發，
後來Microsoft將其作為Windows操作系統的標準協議，並改名為CIFS。
SMB/CIFS是在TCP/IP網絡上運行的，可以讓用戶在不同計算機之間共享文件，
並提供了對共享資源的訪問控制和安全性。

SMBv1/CIFS主要須啟用三項功能: "SMB1Protocol-Client","SMB1Protocol-Server","SMB1Protocol"

在 PowerShell 中，可以使用 DISM 模組來管理和部署 Windows 映像。
該模組提供了一組 cmdlet，用於安裝、卸載、配置和更新 Windows 功能和應用程序。
使用 DISM 模組，您可以查看和修改 Windows 映像文件中的組件、驅動程序和軟件包，
準備系統進行 Windows 安裝之前的操作，以及修復受損的 Windows 系統文件。
此外，DISM 模組還提供了一些 cmdlet 用於管理應用程序和更新，如安裝和卸載應用程序，
檢查和安裝 Windows 更新等。總之，DISM 模組是 PowerShell 中非常有用的模組之一，
可以幫助系統管理員進行各種系統管理和維護任務。
#>

# 20241005 win10和win11的version都是10.開頭
# FIXME: Wiin11無法完成安裝,原因不明待查.

param($runadmin)

function import-module_func ($name) {
    #此function會檢查本機上是否有要載入的模組. 如果沒有, 就連線到wcdc2.vhcy.gov上下載. 可能Win7沒有內建該模組. 
    $result = get-module -ListAvailable $name

    if ($result -ne $null) {

        Import-Module -Name $name -ErrorAction Stop

    }
    else {
        $Username = "vhwcmis"
        $Password = "Mis20190610"
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
        
        $rsession = New-PSSession -ComputerName wcdc2.vhcy.gov.tw -Credential $credential
        Import-Module $name -PSSession $rsession -ErrorAction Stop
        Disconnect-PSSession -Session $rsession | Out-Null
    }
    
}

function Enable-SMBv1 {

    #Windows7預設己安裝.net framework 3, 不需要再裝
       
    $check_win10 = (Get-WmiObject -Class Win32_OperatingSystem | Select-Object -Property Version).Version -like '10*'

    Write-Output "檢查SMBv1/CIFS是否啟用及測試連線172.20.1.14:"

    if ($check_admin -and $check_win10) {

        #載入Dism模組
        import-module_func Dism

        $result = Get-WindowsOptionalFeature -Online -FeatureName "SMB1Protocol"

        if ($result.State -ne "Enabled" ) {
            
            Write-Output "SMBv1/CIFS未啟用,進行啟用SMBv1/CIFS, 等待完成後必須重新開機."

            # 安裝SMB1.0/CIFS功能
            Enable-WindowsOptionalFeature -Online -FeatureName "SMB1Protocol-Client", "SMB1Protocol-Server", "SMB1Protocol" -NoRestart

            # 啟用SMB1.0
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name "SMB1" -Value "1" -Type DWord

            # 啟用CIFS
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" -Name "EnableSecuritySignature" -Value "0" -Type DWord

            # 重啟SMB服務
            Restart-Service -Name "LanmanServer" -force

            # 重啟計算機
            #Restart-Computer -Force -Confirm

            start-sleep -Seconds 5
        }
        else {
        
            Write-Output "SMB己安娤."
        
        }

    }
    else {
        Write-Warning "非Win10或沒有系統管理員權限,無法檢查SMB1.0/CIFS是否有啟用,請以系統管理員身分重新嘗試."
    
    } 
    
}


function Enable-NetFx3 {
    #Windows7預設己安裝.net framework 3, 不需要再裝
    
    #install windowsfeature on win11
    #https://learn.microsoft.com/zh-tw/windows-hardware/manufacture/desktop/enable-net-framework-35-by-using-windows-powershell?view=windows-11
    
    $check_win10 = (Get-WmiObject -Class Win32_OperatingSystem | Select-Object -Property Version).Version -like '10*'

    Write-Output "檢查.NET Framework 3.5(包括.NET2.0和3.0)是否啟用:"
    if ($check_admin -and $check_win10) {

        #載入Dism模組
        import-module_func Dism

        $result = Get-WindowsOptionalFeature -Online -FeatureName "NetFx3"

        if ($result.state -ne "Enabled") {
            Write-Output ".NET Framework 3.5(包括.NET2.0和3.0)未啟,進行啟用, 等待完成後必須重新開機."
            
            $source_path = "\\172.20.5.187\mis\09-dotNet3"

            if (Test-Path -Path $source_path) {

                # 1. 復制檔案到暫存資料夾.
                $source_path = Get-Item $source_path
                copy-item -Path $source_path.FullName -Destination $env:TEMP -Force -Recurse

                # 2. 安裝NETFX3
                Enable-WindowsOptionalFeature -Online -FeatureName "NetFx3", "WCF-HTTP-Activation", "WCF-NonHTTP-Activation" -all -Source "$env:TEMP\$($source_path.name)" -NoRestart

            }
            else {
                Write-Warning "$source_path 找不到, 請檢查路徑."
            }
           
        }
        else {
            Write-Output ".NET Framework 3.5(包括.NET2.0和3.0)己啟用"
        }


    }
    else {
        Write-Warning "沒有系統管理員權限或是Win7系統,無法檢查.NET Framework 3.5(包括.NET2.0和3.0)是否有啟用,請以系統管理員身分重新嘗試."
    }


}



#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Enable-SMBv1
    enable-netfx3
    pause
}