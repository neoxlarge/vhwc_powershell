# 設定Win11 IE mode 企業清單模式, 無時間限制.

param($runadmin)

# 函數：檢查並設置註冊表值
function Set-RegistryValue {
    param (
        [string]$Path,
        [string]$Name,
        $Value,
        [string]$Type
    )

    $params = @{
        Path  = $Path
        Name  = $Name
        Value = $Value
        Type  = $Type
    }

    if (!(Test-Path $Path) ) {
        if ($check_admin) {
            New-Item -Path $Path -Force | Out-Null
        }
        else {
            Write-Output "註冊表路徑不存在: $path, 無管理員權限新增."
        }
    } 

    $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    
    if ($null -eq $currentValue -or $currentValue.$Name -ne $Value) {
        if ($check_admin) {
            Set-ItemProperty @params
            Write-Output "註冊表機碼已設置 $Path\$Name 為 $Value"
        }
        else {
            Write-Output "機碼不正確 $path \ $Name : $($currentValue.Value)"
        }
    }
    else {
        Write-Output "機碼正確 $Path\$Name : $Value"
    }
}

function enable-iemode {

    if ((Get-CimInstance -ClassName Win32_OperatingSystem).Caption -match "Windows 11") {
    
        Write-Output "設定Edge啟用IE模式(企業清單模式):"

        # 主要的註冊表路徑

        $edgePolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
        $popupsAllowedPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"

        # 設置 Edge 策略
        Set-RegistryValue -Path $edgePolicyPath -Name "EnterpriseModeSiteListManagerAllowed" -Value 0 -Type DWord
        Set-RegistryValue -Path $edgePolicyPath -Name "InternetExplorerIntegrationLevel" -Value 1 -Type DWord
        Set-RegistryValue -Path $edgePolicyPath -Name "InternetExplorerIntegrationReloadInIEModeAllowed" -Value 1 -Type DWord
        Set-RegistryValue -Path $edgePolicyPath -Name "InternetExplorerIntegrationSiteList" -Value "\\172.20.1.14\update\0001-中榮系統環境設定\vhwc_win11_IEmode\SiteList.xml" -Type String

        # 設置彈出窗口允許列表
        $popupUrls = @(
            "[*.]vhcy.gov.tw",
            "[*.]vhwc.gov.tw",
            "172.19.[.*][.*]",
            "172.20.[.*][.*]",
            "[*.]vghtc.gov.tw",
            "172.19.[.*][.*]:9090",
            "172.19.[.*][.*]:8000",
            "172.20.[.*][.*]:9090",
            "172.20.[.*][.*]:8000"
        )

        for ($i = 0; $i -lt $popupUrls.Length; $i++) {
            Set-RegistryValue -Path $popupsAllowedPath -Name ($i + 1).ToString() -Value $popupUrls[$i] -Type String
        }

        Write-Output "註冊表設置完成"
    } else {
        Write-Output "系統不是win11, 不使用EDGE IE MODE 企業清單模式."
    }
}


#檔案獨立執行時會執行函式, 如果是被載入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    enable-iemode 

    #pause
    Start-Sleep -Seconds 10
}