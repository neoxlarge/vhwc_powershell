#修改設備控制權, 可該AD的使用者更改印表機設定.

param($runadmin)

function Grant-DeviceOwner {
   

    if ($check_admin) {
        # 定義本機設備擁有者群組名稱和網域使用者群組名稱
        $localGroup = "Device Owner"
        $domainGroup = "vhcy\Domain Users"  

        # 取得本機設備擁有者群組物件
        $localGroupObject = [ADSI]"WinNT://$env:COMPUTERNAME/$localGroup,group"

        # 取得網域使用者群組物件
        $domainGroupObject = [ADSI]"WinNT://$domainGroup,group"

        # 將網域使用者群組新增至本機設備擁有者群組中
        $localGroupObject.Add($domainGroupObject.Path)

        Write-Host "網域使用者群組已新增至設備擁有者群組中。"
    }
    else {
        Write-Warning "沒有系統管理員權限,無法開啟資料?權限,請以系統管理員身分重新嘗試."
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

    Grant-FullControlPermission
    pause
}
