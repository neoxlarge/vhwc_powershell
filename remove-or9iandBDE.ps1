
param($runadmin)

function remove-or9i {

    if ($check_admin) {
        $BDE = Get-WmiObject -class Win32_product | where-object -FilterScript { $_.name -eq "Borland DataBase Engine" }

        if ($BDE -ne $null) {
            $BDE.Uninstall()
        }


        $path = "HKLM:\SOFTWARE\WOW6432Node\ORACLE", 
        "C:\oracle",
        "C:\Program Files (x86)\Oracle",
        "C:\Program Files\Oracle",
        "C:\Program Files (x86)\Common Files\Borland Shared",
        "HKLM:\SOFTWARE\WOW6432Node\Borland"



        foreach ($p in $path) {
            if (Test-Path -Path $p) {
                Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

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

    remove-or9i
    pause
}
