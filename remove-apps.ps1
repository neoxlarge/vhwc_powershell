#移除軟體用(Win10)
param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

#取得OS的版本
function Get-OSVersion {
    $os = (Get-WmiObject -Class Win32_OperatingSystem).Caption

    if ($os -like "*Windows 7*") {
        return "Windows 7"
    }
    elseif ($os -like "*Windows 10*") {
        return "Windows 10"
    }
    else {
        return "Unknown OS"
    }
}


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
        "B9ECED6F.ASUSPCAssistant")              

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


function remove-apps {
    if ($(Get-OSVersion) -eq "Windows 10") {
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
    pause
}


<########################################################################################################

$uninstall_list = @{ name = "onedrive"; version = "0" },
#@{ name = "hicos"; version = "3.0.2" },
@{ name = "skype"; version = "0" }

$all_installed_program = get-installedprogramlist


foreach ($i in $uninstall_list) {

    $app = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "*$($i.name)*" }

    if ($app -ne $null) {
    write-output $app.displayname
    $uninstall_string = $app.UninstallString.Split(" ")
    Write-Output $uninstall_string[2]
    Start-Process -FilePath $uninstall_string[0] -ArgumentList $uninstall_string[2] -wait
    }
   }


#>