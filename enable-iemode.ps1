param($runadmin)


# 函數：檢查並設置 DWORD 值
function Set-RegistryDWORD {
    param (
        [string]$Path,
        [string]$Name,
        [int]$Value
    )
    if (!(Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($currentValue -eq $null -or $currentValue.$Name -ne $Value) {
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type DWord
        Write-Output "已設置 $Path\$Name 為 $Value"
    } else {
        Write-Output "$Path\$Name 已存在且值正確"
    }
}

# 函數：檢查並設置字符串值
function Set-RegistryString {
    param (
        [string]$Path,
        [string]$Name,
        [string]$Value
    )
    if (!(Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($currentValue -eq $null -or $currentValue.$Name -ne $Value) {
        Set-ItemProperty -Path $Path -Name $Name -Value $Value
        Write-Output "已設置 $Path\$Name 為 $Value"
    } else {
        Write-Output "$Path\$Name 已存在且值正確"
    }
}

function enable-iemode {

# 主要的註冊表路徑

$edgePolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
$popupsAllowedPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"

# 設置 Edge 策略
Set-RegistryDWORD -Path $edgePolicyPath -Name "EnterpriseModeSiteListManagerAllowed" -Value 0
Set-RegistryDWORD -Path $edgePolicyPath -Name "InternetExplorerIntegrationLevel" -Value 1
Set-RegistryDWORD -Path $edgePolicyPath -Name "InternetExplorerIntegrationReloadInIEModeAllowed" -Value 1
Set-RegistryString -Path $edgePolicyPath -Name "InternetExplorerIntegrationSiteList" -Value "\\172.20.1.14\update\0001-中榮系統環境設定\vhwc_win11_IEmode\SiteList.xml"

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
    Set-RegistryString -Path $popupsAllowedPath -Name ($i + 1).ToString() -Value $popupUrls[$i]
}

Write-Output "註冊表設置完成"
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