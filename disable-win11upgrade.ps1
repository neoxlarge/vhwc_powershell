#修正

param($runadmin)

#取得OS的版本
function Get-OSVersion {
    $os = (Get-WmiObject -Class Win32_OperatingSystem).Caption

    if ($os -like "*Windows 7*") {
        return "Windows 7"
    }
    elseif ($os -like "*Windows 10*") {
        return "Windows 10"
    }
    elseif ($os -like "*Windows 11*") {
        return "Windows 11"
    }
    else {
        return "Unknown OS"
    }
}

function disable-win11upgrade {
    
    #Only Win10 need to change the registry.

    if ((Get-OSVersion) -eq "Windows 10") {

        set-itemproperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -name SvOfferDeclined -value 1697695220682 -type QWord

    } 
    else {
    Write-Output "Not Win 10, no need to disable update to Win 11."
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

    disable-win11upgrade
    

    pause
}
