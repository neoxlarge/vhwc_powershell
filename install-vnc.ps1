# 安裝VNC
# setup.exe /? 可以顯示安裝指令

# C:\ProgramData\Microsoft\Windows\Start Menu\Programs\UltraVNC\UltraVNC Server.lnk
# C:\ProgramData\Microsoft\Windows\Start Menu\Programs\UltraVNC\UltraVNC Viewer.lnk


param($runadmin)

$log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VHWC_logs\update-VNC.log"

function import-vhwcmis_module {
    $moudle_paths = @(
        if ($script:MyInvocation.MyCommand.Path) {"$(Split-Path $script:MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue)"},
        "d:\mis\vhwcmis",
        "c:\mis\vhwcmis",
        "\\172.20.5.185\powershell\vhwc_powershell",
        "\\172.20.1.14\share\00新電腦安裝\vhwc_powershell",
        "\\172.20.1.122\share\software\00newpc\vhwc_powershell",
        "\\172.19.1.229\cch-share\h040_張義明\vhwc_powershell"
    )

    $filename = "vhwcmis_module.psm1"

    foreach ($path in $moudle_paths) {
        
        if (Test-Path "$path\$filename") {
            write-output "$path\$filename"
            Import-Module "$path\$filename" -ErrorVariable $err_import_module
            if ($err_import_module -eq $null) {
                Write-Output "Imported module path successed: $path\$filename"
                break
            }
        }
    }

    $result = get-command -Name "get-installedprogramlist" -CommandType Function -ErrorAction SilentlyContinue
    if ($result -eq $null) {
        throw "無法載入vhwcmis_module模組, 程式無法正常執行. "
    }
}
import-vhwcmis_module


# 定義創建捷徑的函數
function New-Shortcut {
    param (
        [string]$TargetPath,
        [string]$ShortcutPath
    )
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.Save()
}

function install-VNC {
    
    $software_name = "UltraVnc"
    $software_path = "\\172.20.5.187\mis\08-VNC\1_4_36"
    # FIXME: $software_path = "\\172.20.1.122\share\software\00newpc\08-VNC\1_4_36"
    $software_msi = "UltraVNC_X64.msi"
    $software_msi_x86 = "UltraVNC_X86.msi"
    $log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VHWC_logs\update-VNC.log"

    ## 判斷OS是32(x86)或是64(AMD64), 其他值(ARM64)不安裝  
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $software_exec = $software_msi }
        "x86" { $software_exec = $software_msi_x86 }
        default { Write-Warning "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
    }

    ## 找出軟體是否己安裝

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }


    if ($software_is_installed) {

        $msi_version = get-msiversion -MSIPATH "$software_path\$software_exec"
        $result = Compare-Version -Version1 $msi_version -Version2 $software_is_installed.DisplayVersion

        if ($result) {
            $ipv4 = Get-IPv4Address 
            Write-Log -logfile $log_file -message "Find old VNC version:$($software_is_installed.DisplayVersion)"
            Write-Output "Find old VNC $software_name, version: $($software_is_installed.DisplayVersion)"
            Write-Output "Removing old version."
            Start-Process -FilePath $software_is_installed.UninstallString -ArgumentList "/SILENT" -Wait
            $software_is_installed = $null

        } else {
            $msg = "Installed VNC: $($software_is_installed.DisplayVersion)"
            #Write-Log -logfile $log_file -message $msg
        
        }
    }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #復制檔案到本機暫存"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        <#
        SERVERVIEWER
        1 server
        2 viewer
        3 server + viewer
        SERVICE
        1 install
        2 not install

        PASSWORD = mypassword
        Sample
        UltraVNC_1436_X86.msi  SERVERVIEWER=3 SERVICE=1 PASSWORD="sysc0012"
        #>

        if ($null -ne $software_exec) {
            #Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/silent /loadinf=$($env:temp + "\" + $software_path.Name + "\installvnc.inf")" -Wait
            $install_filepath = "$($env:temp)\$($software_path.Name)\$software_exec"
            $install_arrg = "/i $install_filepath /passive /norestart PASSWORD=""sysc0012"" SERVERVIEWER=3 SERVICE=1"
            Start-Process -FilePath "msiexec.exe" -ArgumentList $install_arrg -Wait
          
            Write-Log -logfile $log_file -message "Start to install VNC: $install_filepath"

            Start-Sleep -Seconds 5
        }
        else {
            Write-Warning "$software_name 無法正常安裝."
        }
      
        #復制設定檔vltravnc.ini 到C:\Program Files\uvnc bvba\UltraVNC
        Copy-Item -Path ($env:temp + "\" + $software_path.Name + "\ultravnc.ini") -Destination ($env:ProgramFiles + "\uvnc bvba\UltraVNC") -Force

        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    
    }

    #因為這版用MSI安裝的, 會出現沒有建立程式捷徑的情, 所以手動多建立一個公用桌面的捷徑

    $shortcuts = @{
        
        "viewer" = @{
            "name" = "UltraVNC Viewer.lnk"
            "folder" = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\UltraVNC"
            "exe" = "C:\Program Files\uvnc bvba\UltraVNC\vncviewer.exe"
        }

        "server" = @{
            "name" = "UltraVNC Server.lnk"
            "folder" = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\UltraVNC"
            "exe" = "C:\Program Files\uvnc bvba\UltraVNC\winvnc.exe"
        }
    
    }

    foreach ($shortcut in $shortcuts.keys) {

        $folderPath = $shortcuts.$shortcut.folder
        $shortcutPath = Join-Path $folderPath $shortcuts.$shortcut.name
        $exePath = $shortcuts.$shortcut.exe


        # 創建資料夾（如果不存在）
        if (!(Test-Path -Path $folderPath)) {
            try {
                New-Item -Path $folderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            catch {
                Write-Log -logfile $log_file -message "錯誤：無法創建資料夾 $folderPath - $_"
                continue
            }
        }


        # 創建捷徑

        if (!(Test-Path -Path $shortcutPath)) {
            try {
                New-Shortcut -TargetPath $exePath -ShortcutPath $shortcutPath
                Write-Log -logfile $log_file -message "成功創建捷徑：$shortcutPath"
            }
            catch {
                Write-Log -logfile $log_file -message "錯誤：無法創建捷徑 $shortcutPath - $_"
            }
         }
    }
    
  
    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)

}


#檔案獨立執行時會執行函式, 如果是被載入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    if ($check_admin) { 
        install-VNC    
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    #pause
    Start-Sleep -Seconds 10
}


