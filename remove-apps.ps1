#移除軟體用(Win10)
param($runadmin)

function import-vhwcmis_module {
    $moudle_paths = @(
        if ($script:MyInvocation.MyCommand.Path) {"$(Split-Path $script:MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue)"},
        "d:\mis\vhwcmis",
        "c:\mis\vhwcmis",
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


function Remove-AppsinWin10 {
    # 要移除的應用程式清單 (win10 only)
    $appsToRemove = @(
        "Microsoft.SkypeApp", 
        "Microsoft.OneDrive", 
        "Microsoft.MicrosoftOfficeHub",
        #"Microsoft.XboxIdentityProvider",   #此應用程式是 Windows 的一部分，而且無法針對個別使用者解除安裝。
        #"Microsoft.XboxGameCallableUI",     #此應用程式是 Windows 的一部分，而且無法針對個別使用者解除安裝。
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.XboxGamingOverlay",      
        "Microsoft.XboxGameOverlay",        
        "Microsoft.XboxApp",                
        "Microsoft.Xbox.TCUI",
        "5A894077.McAfeeSecurity",
        "B9ECED6F.ASUSPCAssistant",
        "4DF9E0F8.Netflix",
        "ZhuhaiKingsoftOfficeSoftw.WPSOffice")              

    # 遍歷每個要移除的應用程式
    foreach ($app in $appsToRemove) {
        # 檢查是否已經安裝該應用程式
        if (Get-AppxPackage -Name $app) {
            # 移除應用程式
            Remove-AppxPackage -Package $(Get-AppxPackage -Name $app)
            Write-output "已成功移除 $app 應用程式。"
        }
        else {
            Write-output "$app 應用程式未安裝。"
        }
    }
}


function update-apps {

    # 設定 $DebugPreference 為 "Continue" 顯示訊息.
    $DebugPreference = "Continue"
    $apps = @{

        "StickyNotes" = @{
            "name" = "Microsoft.MicrosoftStickyNotes"
            "version" = "6.1.2.0"
            "path" = "\\172.20.1.122\share\software\00newpc\40-Microsoft_Store"
            "filename" = "microsoft-sticky-notes-6-1-2-0.msixbundle"
        }
        #　Photos有OS版本要求.
        "Photos" = @{
            "name" = "Microsoft.Windows.Photos"
            "version" = "2024.11070.31001.0"
            "path" = "\\172.20.1.122\share\software\00newpc\40-Microsoft_Store"
            "filename" = "microsoft-photos-2024-11070-31001-0.msixbundle"
        }
    }

    foreach ($app in $apps.Keys) {
        Write-Debug "Check appx software (msixbundle): $($apps.$app.name) version: $($apps.$app.version)"
        $installed_version = (Get-AppxPackage -Name $apps.$app.name).Version
        Write-Debug "Installed version: $installed_version"
        $result = Compare-Version -Version1 $apps.$app.version -Version2 $installed_version

        if ($result) {
            Write-Output "更新Appx: $($apps.$app.name)"
            $app_fullpath = "$($apps.$app.path)\$($apps.$app.filename)"
            Write-Debug "App fullpath = $($app_fullpath)"
            Add-AppxPackage -Path $app_fullpath -Update
        } else {
            Write-Output "Appx: $($apps.$app.name) 不更新."
        }
    }

}


function remove-apps {
    if ($(Get-OSVersion) -in "Windows 10","Windows 11") {
        Remove-AppsinWin10
    }
}

#檔案獨立執行時會執行函式, 如果是被匯入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    remove-apps
    update-apps
    pause
}

