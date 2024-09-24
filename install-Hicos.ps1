# 安裝Hicos
# 輸入setup.exe /?, 可以看到支援silent安裝的指命
# /install /repair /uninstall 
# /passive /quiet
# /norestart /log log.txt


param($runadmin,[switch]$gpo)

$DebugPreference = "Continue"
$log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VHWC_logs\install-Hicos.log"

function import-vhwcmis_module {
    $moudle_paths = @(
        if ($script:MyInvocation.MyCommand.Path) {"$(Split-Path $script:MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue)"},
        "d:\mis\vhwcmis",
        "c:\mis\vhwcmis",
        #從網路上import的,可能會被擋下.
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
                Write-Output "Imported module path successed: $path\$filename"
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

function install-HiCOS {


    $software_name = "HiCOS*"
    $software_path = "\\172.20.5.187\mis\04-自然人憑證-新版HiCOS\HiCOS_Client-3.1.0.22133-20220624"
    $software_exec = "HiCOS_Client.exe"
    $software_version = get-itemproperty -Path "$software_path\$software_exec" -Name VersionInfo -ErrorAction SilentlyContinue
    $software_version = $software_version.VersionInfo.ProductVersion

    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    ## 公文系統會安裝2.1.9.1, 移掉這版

    foreach ($software in $software_is_installed) {
        if (Compare-Version -Version1 $software_version  -Version2 $software.DisplayVersion) {
            Write-Output "Find old $software_name version: $software_version, uninstall it."
            if ($gpo) {Write-Log -LogFile $log_file -Message "Find old $software_name version: $($software.DisplayVersion), uninstall it."}
            
            #$uninstallstring = $software.uninstallString.Split(" ")[1].replace("/I", "/x")
            #Start-Process -FilePath "msiexec.exe" -ArgumentList "$uninstallstring /passive" -Wait
            #Start-Sleep -Seconds 5
            
            Uninstall-Software -Name $software.DisplayName
            
            $software_is_installed = $software_is_installed | Where-Object { $_.DisplayVersion -ne $software.DisplayVersion } #從陣列中移除
        }
    }


    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到本機暫存"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        $pid_ = Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/install /passive" -PassThru #不知為什麼-wait會停不下來.
        while (!($pid_.HasExited)) { Start-Sleep -Seconds 5}
        
        if ($gpo) {Write-Log -logfile $log_file -message "Install $software_name  version: $software_version"}
        
        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

        #change HiCOS setting
        
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" {
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card V3 (Conf.1 T=CL)" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card V3 (Conf.1 T=1)" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card 32K" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Cryptography\Calais\SmartCards\CHT ePKI Card 32K" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Chunghwa TeleCom\HiCOS PKI Smart Card\TokenUtility\1.0.0" -Name "CertImportWithLegacyCSP" -Value 1 -Force -PropertyType "DWord"
                Set-ItemProperty -Path "HKLM:\Software\WOW6432Node\Chunghwa TeleCom\HiCOS PKI Smart Card\TokenUtility\1.0.0" -Name "LegacyCSPParing" -Value 1 -Force
            }
            "x86" {
                New-ItemProperty -Path "HKLM:\Software\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card V3 (Conf.1 T=CL)" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card V3 (Conf.1 T=1)" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card 32K" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\Microsoft\Cryptography\Calais\SmartCards\CHT ePKI Card 32K" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\Chunghwa TeleCom\HiCOS PKI Smart Card\TokenUtility\1.0.0" -Name "CertImportWithLegacyCSP" -Value 1 -Force -PropertyType "DWord"
                Set-ItemProperty -Path "HKLM:\Software\Chunghwa TeleCom\HiCOS PKI Smart Card\TokenUtility\1.0.0" -Name "LegacyCSPParing" -Value 1 -Force
            }
        }
    
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
        install-HiCOS
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}