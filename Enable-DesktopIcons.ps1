#桌面環境設定
#啟用桌面圖示

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
    else {
        return "Unknown OS"
    }
}

function Enable-DesktopIcons {
    Write-host "設定桌面 ."
    #Win10和Win7都相同的設定值改法

    if (!(Test-Path -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel")) {
        New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Force | out-null
    }

    #電腦 
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value 0
    #使用者文件圖示
    #Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{59031a47-3f72-44a7-89c5-5595fe6b30ee}" -Value 0 
    #網路
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}" -Value 0 
    #控制台
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}" -Value 0
    #資源回收筒
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" -Name "{645FF040-5081-101B-9F08-00AA002F954E}" -Value 0

    if ((Get-OSVersion) -eq "Windows 10") {
        #Windows 10 才需要執行這

        # 啟用桌面圖示 (Win10)
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideIcons" -Value 0
        # 隱藏 Cortana 按鈕(Win10)
        Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowCortanaButton" -Value 0
        # 隱藏 搜尋 按鈕(Win10)
        Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0
        # 隱藏 新聞和興趣 按鈕(Win10)
        # 參考https://admx.help/?Category=Windows_10_2016&Policy=Microsoft.Policies.Feeds::EnableFeeds
        if ($check_admin) {
            #寫入Hotkey local machine 需要管理員權限.
            $reg_path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds"
            if (Test-Path -Path $reg_path) {
                Set-ItemProperty -Path $reg_path -Name "EnableFeeds" -Value 0 -Force
            }
            else {
                New-Item -Path $reg_path -force
                New-ItemProperty -Path $reg_path -Name "EnableFeeds" -Value 0 -PropertyType DWord -force
            }
        }
    }

    #重啟桌面
    Stop-Process -Name "Explorer" -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
    
    #取消啟動時執行OneDrive.
    $result = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" 
    if ($result.OneDrive -ne $null) {
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "OneDrive" -Value $null
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

    Enable-DesktopIcons
    
    pause
}