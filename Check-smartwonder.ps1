<#check-smartwonder
SmartWonder是用瀏覽器看PACS的
1. 開敫smartwonder會要求寫入c:\program files\tedpc
2. 會下載安裝2個元件, 但憑證己過期, 要改IE選項.
3. 在非PACS允許的電腦上, 會開眼睛來看PACE, ultraquery 有2個選項要改.
4. IE和EDGE設定方不同,EDGE也是要用IE mode.

#>

param($runadmin)

#管理者權限vhwcmis的證書.
$Username = "vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)


function check_smartwonder {

    # 1. check c:\program files\tedpc
    if (!(Test-Path -path "c:\program files\tedpc")) {
        New-Item -Path "c:\program files\tedpc" -ItemType Directory -Force
    }
    
    # grant control to the folder
    # 留到grant-fullcontrolpermission再統一開權限.

    # 2.會下載安裝2個元件, 但憑證己過期, 要改IE選項.
    <#
    即使簽章無效也允許執行或安裝軟體
    這個原則設定可讓您管理即使簽章無效時，使用者是否能安裝或執行諸如 ActiveX 控制項和檔案下載的軟體。無效的簽章可能表示已有人竄改檔案。
    如果您啟用這個原則設定，將會在安裝或執行具有無效簽章的檔案時提示使用者。
    如果您停用這個原則設定，使用者將無法安裝或執行具有無效簽章的檔案。
    如果您未設定這個原則，使用者可以選擇安裝或執行具有無效簽章的檔案。
    支援的作業系統: 在 Windows XP 含 Service Pack 2 或 Windows Server 2003 含 Service Pack 1 至少需要 Internet Explorer 6.0

    Registry Hive	HKEY_LOCAL_MACHINE or HKEY_CURRENT_USER
    Registry Path	Software\Policies\Microsoft\Internet Explorer\Download
    Value Name	RunInvalidSignatures
    Value Type	REG_DWORD
    Enabled Value	1
    Disabled Value	0
    #>
    Set-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Internet Explorer\Download" -Name "RunInvalidSignatures" -Value 1 -Force

    


}

#檔案獨立執行時會執行函式, 如果是被滙入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    Check-VncSetting
    Check-VncService

    pause
}
