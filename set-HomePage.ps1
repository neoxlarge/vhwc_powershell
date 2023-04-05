#修改三個瀏覽器的預設首頁, IE, Edge, chrome.
#20230329, 網域一般使用者無法寫入HKCU:\SOFTWARE\Policies, 改成由管理者寫到HKLM:\SOFTWARE\Policies
param($runadmin)

function Set-HomePage {
    if ($check_admin) {
        $HomePage = "https://eip.vghtc.gov.tw"
        Write-Output "設定IE,Edage,Chrome 預設開啟首頁為: $HomePage"
        # Modify Edge home page
        # 參考連結: https://learn.microsoft.com/en-us/deployedge/microsoft-edge-policies#restoreonstartup
        $reg_path = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\RestoreOnStartupURLs"
        $result = Test-Path -Path $reg_path
        if ($result -eq $false) {
            New-Item -Path $reg_path -force
        }
        Set-ItemProperty -Path $reg_path -Name "1" -Value $HomePage
        #Edge可以設定多個，name值是數字一直加.
        #Set-ItemProperty -Path $reg_path -Name "2" -Value "https://www.vghtc.gov.tw"
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "RestoreOnStartup" -Value 4

        # Modify Chrome home page
        $reg_path = "HKLM:\Software\Policies\Google\Chrome\RestoreOnStartupURLs"
        $result = Test-Path -Path $reg_path
        if ($result -eq $false) {
            New-Item -Path $reg_path -force
        }
        Set-ItemProperty -Path $reg_path -Name "1" -Value $HomePage
        Set-ItemProperty -Path "HKLM:\Software\Policies\Google\Chrome" -Name "RestoreOnStartup" -Value 4


        # Modify Internet Explorer home page
        Set-ItemProperty -Path "HKLM:\Software\Microsoft\Internet Explorer\Main" -Name "Start Page" -Value $HomePage
        Set-ItemProperty -Path "HKLM:\Software\Microsoft\Internet Explorer\Main" -Name "Default_Page_URL" -Value $HomePage

    } else {
        Write-Warning "沒有系統管理員權限,無法強制設定Homepage,請以系統管理員身分重新嘗試。"
    }
}




#檔案獨立執行時會執行函式, 如果是被?入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    Set-HomePage
    pause
}
