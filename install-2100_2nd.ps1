# 新的公文系統 20240729上線
# 系統需求:
# HCAserversign 醫事人員卡用
# Hiicos 自然人用
#  desktop放捷徑, 看chrome安裝位置
# 機碼, 用於彈出式window

param($runadmin)

$log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VHWC_logs\install-2100_2nd.log"

function import-vhwcmis_module {
    # import vhwcmis_module.psm1
    # 取得vhwcmis_module.psm1的3種方式:
    # 1.程式執行當前路徑, 放到AD上用Group police執行,不會有當前路徑.
    # 2.常用的路徑, d:\mis\vhwc_powershell, 不是每台都有放.
    # 3.連到NAS上取得. 非網域的電腦會沒有NAS的權限, 須手動連上NAS.

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
}
import-vhwcmis_module



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
            
    # 檢查並創建註冊表項（如果不存在）
    if (-not (Test-Path $keypath)) {
        $createItemCode = {
            param($path)
            New-Item -Path $path -Force
        }

        $scriptString = $createItemCode.ToString()
        $argumentList = "-ExecutionPolicy Bypass -Command `"& {$scriptString} -path '$keypath'`""

        $proc =Start-Process powershell.exe -ArgumentList $argumentList -Credential $credential -PassThru
        $proc.WaitForExit()
        Write-Output "建立新的註冊表項: $keypath"
    }

    # 檢查註冊表值是否存在且正確
        $currentValue = Get-ItemProperty -Path $keypath -Name $valuename -ErrorAction SilentlyContinue
        if ($currentValue -eq $null -or $currentValue.$valuename -ne $desiredValue) {
            $needsUpdate = $true
        } else {
            $needsUpdate = $false
        }
   
   
    if ($needsUpdate) {
        # 更新註冊表值
        $updatePropertyCode = {
            param($path, $name, $value)
            New-ItemProperty -Path $path -Name $name -Value $value -PropertyType String -Force
        }

        $scriptString = $updatePropertyCode.ToString()
        $argumentList = "-ExecutionPolicy Bypass -Command `"& {$scriptString} -path '$keypath' -name '$valueName' -value '$desiredValue'`""
        $proc = Start-Process powershell.exe -ArgumentList $argumentList -Credential $credential -PassThru
        $proc.WaitForExit()
       
        Write-Output "二代公文系統更新註冊表項: $keypath\$valueName"
        Write-Log -LogFile $log_file -Message "二代公文系統更新註冊表項: $keypath\$valueName"
    }
    else {
        Write-Output "二代公文系統註冊表項已存在且正確: $keypath\$valueName"
    }
}



function install-2100_2nd() {

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
            #Write-Output "二代公文系統捷徑已存在，但 Chrome 路徑不正確。正在更新..."
            #$shortcut.TargetPath = $chromePath
            #$shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
            #$shortcut.Save()
            
            $shortcutPath_temp = Join-Path $env:temp "二代公文系統(Chrome).lnk"

            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcutPath_temp)
            $Shortcut.TargetPath = $chromePath
            $Shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
            $Shortcut.Save()

            $credential = get-admin_cred
            Start-Process -FilePath robocopy.exe -ArgumentList "$($env:temp) C:\Users\Public\Desktop 二代公文系統(Chrome).lnk" -Credential $credential
            
            
            Write-Output "二代公文系統捷徑已更新。"
            Write-log -LogFile $log_file -Message "原有二代公文系統捷徑內容有誤,捷徑已更新。  "
        }
        else {
            Write-Output "二代公文系統捷徑已存在且 Chrome 路徑正確。無需更改。"
        }
    }
    else {
        # 建立新捷徑
        $shortcutPath_temp = Join-Path $env:temp "二代公文系統(Chrome).lnk"

        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath_temp)
        $Shortcut.TargetPath = $chromePath
        $Shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
        $Shortcut.Save()

        $credential = get-admin_cred
        Start-Process -FilePath robocopy.exe -ArgumentList "$($env:temp) C:\Users\Public\Desktop 二代公文系統(Chrome).lnk" -Credential $credential
        Write-Output "新捷徑 '二代公文系統(Chrome).lnk' 已建立完成。"
        Write-log -LogFile $log_file -Message "新捷徑 '二代公文系統(Chrome).lnk' 已建立完成。"
    }

    # 檢查和更新註冊表項
    $chromeKeyPath = "HKLM:\SOFTWARE\Policies\Google\Chrome\PopupsAllowedForUrls"
    $edgeKeyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"
    $desiredValue = "edap.doc.vghtc.gov.tw"


    Update-RegistryKey $chromeKeyPath "99999" $desiredValue
    Update-RegistryKey $edgeKeyPath "99999" $desiredValue

    Write-Output "二代公文系統註冊表檢查和更新完成。"

}





#檔案獨立執行時會執行函式, 如果是被載入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    install-2100_2nd    

    #pause
    Start-Sleep -Seconds 10
}