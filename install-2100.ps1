# 新的公文系統 20240729上線
# 系統需求:
# HCAserversign 醫事人員卡用
# Hiicos 自然人用
#  desktop放捷徑, 看chrome安裝位置
# 機碼, 用於彈出式window

param($runadmin)

$log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VHWC_logs\install-2100(chrome).log"

$pspaths = @()

if ($script:MyInvocation.MyCommand.Path -ne $null) {
    $work_path = "$(Split-Path $script:MyInvocation.MyCommand.Path)\vhwcmis_module.psm1"
    if (test-path -Path $work_path) { $pspaths += $work_path }
}
$nas_name = "nas122"
$nas_path = "\\172.20.1.122\share\software\00newpc\vhwc_powershell"
if (!(test-path $nas_path)) {
    $nas_Username = "software_download"
    $nas_Password = "Us2791072"
    $nas_securePassword = ConvertTo-SecureString $nas_Password -AsPlainText -Force
    $nas_credential = New-Object System.Management.Automation.PSCredential($nas_Username, $nas_securePassword)
    
    New-PSDrive -Name $nas_name -Root "$nas_path" -PSProvider FileSystem -Credential $nas_credential | Out-Null
}
$pspaths += "$nas_path\vhwcmis_module.psm1"

$local_path = "d:\mis\vhwc_powershell\vhwcmis_module.psm1"
if (Test-Path $local_path) { $pspaths += $local_path }

foreach ($path in $pspaths) {
    Import-Module $path -ErrorAction SilentlyContinue
    if ((get-command -Name "get-installedprogramlist" -CommandType Function -ErrorAction SilentlyContinue)) {
        break
    }
}



# 函數：獲取 Chrome 的實際安裝路徑
function Get-ChromePath {
    $chromePaths = @(
        (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"),
        (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe")
    )

    foreach ($path in $chromePaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    Write-Output "無法找到 Chrome。請確保 Chrome 已安裝。"
    exit
}

function Update-RegistryKey($keyPath, $valueName, $desiredValue) {
    # 創建一個新的 PSDrive 來訪問 HKLM
    New-PSDrive -Name HKLM -PSProvider Registry -Root HKEY_LOCAL_MACHINE -ErrorAction SilentlyContinue | Out-Null

    $fullPath = "HKLM:\$keyPath"
    if (-not (Test-Path $fullPath)) {
        New-Item -Path $fullPath -Force | Out-Null
        Write-Output "建立新的註冊表項: $fullPath"
    }

    $currentValue = Get-ItemProperty -Path $fullPath -Name $valueName -ErrorAction SilentlyContinue
    if ($currentValue -eq $null -or $currentValue.$valueName -ne $desiredValue) {
        New-ItemProperty -Path $fullPath -Name $valueName -Value $desiredValue -PropertyType String -Force | Out-Null
        Write-Output "更新註冊表項: $fullPath\$valueName"
    }
    else {
        Write-Output "註冊表項已存在且正確: $fullPath\$valueName"
    }

    # 移除 PSDrive
    Remove-PSDrive -Name HKLM -ErrorAction SilentlyContinue
}



function install-2100_chrome() {

    $credential = get-admin_cred

    # 函數：獲取 Chrome 的實際安裝路徑


    # 獲取 Chrome 路徑
    $chromePath = Get-ChromePath

    # 設定捷徑路徑
    $shortcutPath = Join-Path ([System.Environment]::GetFolderPath("CommonDesktopDirectory")) "二代公文系統(Chrome).lnk"

    # 檢查捷徑是否已存在
    if (Test-Path $shortcutPath) {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
    
        if ($shortcut.TargetPath -ne $chromePath) {
            Write-Output "捷徑已存在，但 Chrome 路徑不正確。正在更新..."
            $shortcut.TargetPath = $chromePath
            $shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
            $shortcut.Save()
            Write-Output "捷徑已更新。"
        }
        else {
            Write-Output "捷徑已存在且 Chrome 路徑正確。無需更改。"
        }
    }
    else {
        # 建立新捷徑
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = $chromePath
        $Shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
        $Shortcut.Save()

        Write-Output "新捷徑 '二代公文系統(Chrome).lnk' 已建立完成。"
    }

    # 檢查和更新註冊表項
    $chromeKeyPath = "SOFTWARE\Policies\Google\Chrome\PopupsAllowedForUrls"
    $edgeKeyPath = "SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"
    $desiredValue = "edap.doc.vghtc.gov.tw"



    Update-RegistryKey $chromeKeyPath "99999" $desiredValue
    Update-RegistryKey $edgeKeyPath "99999" $desiredValue

    Write-Output "註冊表檢查和更新完成。"

}





#檔案獨立執行時會執行函式, 如果是被載入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    install-2100_chrome    

    #pause
    Start-Sleep -Seconds 10
}