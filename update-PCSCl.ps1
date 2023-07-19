# 升級雲端安控元件健保卡讀卡機控制(PCSC 5.1.5.7)

param($runadmin)


$opd_table = @{
    "opd101" = @{
        "no" = "101";
        "name" = "wnur-opd-pc01"; 
        "ip" = "172.20.9.1"
    };
    "opd103" = @{
        "no" = "103";
        "name" = "wnur-pod-pc02";
        "ip" = "172.20.9.2"
    };
    "opd105" = @{};
    "opd106" = @{
        "no" = "106";
        "name" = "wnur-opd-pc04";
        "ip" = "172.20.9.4"
    };
    "opd107" = @{
        "no" = "107";
        "name" = "wnur-opd-pc07";
        "ip" = "172.20.9.7"
    };
    "opd108" = @{
        "no" = "108";
        "name" = "wnur-opd-pc06";
        "ip" = "172.20.9.6"
    };
    "opd109" = @{
        "no" = "109";
        "name" = "wnur-opd-pc05";
        "ip" = "172.20.9.5"
    };
    "opd201" = @{
        "no" = "201";
        "name" = "wnur-opd-pc23";
        "ip" = "172.20.12.23"
    };
    "opd205" = @{};
    "opd206" = @{};
    "erx" = @{};
    "wreh" = @{             #復建科
        "no" = "wreh";
        "name" = "wrch-000-pc01";
        "ip" = "172.20.17.61"
    }

}

foreach ($i in $opd_table.keys) {
    Write-Output $opd_table.$i.name
}


function Get-IPv4Address {
    <#
    回傳找到的IP,只能在172.*才能用. 
    #>

    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.20.*" } |
    Select-Object -ExpandProperty IPAddress |
    Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
    Select-Object -First 1

    if ($ip -eq $null) {
        return $null
    }
    else {     
        return $ip
    }
}

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

function Check-EnvPathContains($path) {
    # 檢查系統環境變數$env:path是否包含特定路徑
    $envPathList = $env:path -split ";"
    foreach ($p in $envPathList) {
        if ($p -like "*$path*") {
            #系統環境變數中包含$path,不需執行C:\VGHTC\00_mis\中榮iccard環境變數設定.bat 
            return $true
        }
    }
    #系統環境變數中不包含$path,需執行C:\VGHTC\00_mis\中榮iccard環境變數設定.bat 
    return $false
}


function Create-Shortcut {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath,

        [Parameter(Mandatory = $true)]
        [string]$ShortcutPath
    )

    # 創建 WScript.Shell 物件
    $shell = New-Object -ComObject WScript.Shell

    # 檢查快捷方式是否已存在，如果存在勍刪除
    if ($shell.CreateShortcut($ShortcutPath).FullName) {
        Remove-Item $ShortcutPath -Force -ErrorAction SilentlyContinue
    }

    # 創建快捷
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Save()
    
}

function get-installedprogramlist {
    # 取得所有安裝的軟體,底下安裝軟體會用到.

    ### Win32_product的清單並不完整， Winnexus 並不在裡面.
    ### $all_installed_program = Get-WmiObject -Class Win32_Product

    ### 所有的軟體會在底下這三個登錄檔路徑中

    $software_reg_path = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    return (Get-ItemProperty -Path $software_reg_path -ErrorAction SilentlyContinue)
}


function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI",                 # Line Notify 存取權杖

        [Parameter(Mandatory = $true)]
        [string]$Message,               # 要發送的訊息內容

        [string]$StickerPackageId,      # 要一併傳送的貼圖套件 ID

        [string]$StickerId              # 要一併傳送的貼圖 ID
    )

    # Line Notify API 的 URI
    $uri = "https://notify-api.line.me/api/notify"

    # 設定 HTTP Header，包含 Line Notify 存取權杖
    $headers = @{ "Authorization" = "Bearer $Token" }

    # 設定要傳送的訊息內容
    $payload = @{
        "message" = $Message
    }

    # 如果要傳送貼圖，加入貼圖套件 ID 和貼圖 ID
    if ($StickerPackageId -and $StickerId) {
        $payload["stickerPackageId"] = $StickerPackageId
        $payload["stickerId"] = $StickerId
    }

    try {
        # 使用 Invoke-RestMethod 傳送 HTTP POST 請求
        Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payload

        # 訊息成功傳送
        Write-Output "訊息已成功傳送。"
    }
    catch {
        # 發生錯誤，輸出錯誤訊息
        Write-Error $_.Exception.Message
    }
}


function update-pcsc {
    Write-Host "升級健保卡讀卡機控制軟體PCSC 5.1.5.7"
    # 軟體環境檢查
    # 1. 取得己安裝版本
    $installedPCSC = Get-WmiObject -Class Win32_Product | Where-Object -FilterScript { $_.name -like "*PCSC*" }
    
    # 2. 取得ip
    $ipv4 = Get-IPv4Address 

    #條件:
    #1安裝的版本為5.1.5.3 或 51.5.5
    #2符合限制的IP

    #此段為powershell v2語法, v2 不支援-in語法
    $installedVersions = "5.1.51","5.1.53", "5.1.55"
    $check_version = $false

    foreach ($version in $installedVersions) {
        if ($installedPCSC.version -eq $version) {
            $check_version = $true
            break
        }
    }

    $check_condition = $check_version -and ($ipv4 -like "172.20.*")


    #寫入記錄檔路徑
    $log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\pcsc_5157.log"
    
    $log_string = "Boot,PCSC:$($installedPCSC.version),$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)"
    $log_string | Add-Content -PassThru $log_file
    Send-LineNotifyMessage -Message $log_string
    Start-Sleep -Seconds 2
    
    #符合升級條件,進行升級
    if ($check_condition) {
                
        #復制新版
        $new_pcsc_path = "\\172.20.5.187\mis\23-讀卡機控制軟體\CMS_CS5.1.5.7_20220925\CS5.1.5.7版_20220925"
    
        $new_pcsc_path = Get-Item $new_pcsc_path
        Copy-Item -Path $new_pcsc_path -Destination "C:\Vghtc\00_mis\$($new_pcsc_path.name)" -Recurse -Force -Verbose

        #關閉健保卡軟體
        Stop-Process -Name csfsim -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3

        #移除舊版
        $installedPCSC.uninstall()
        Start-Sleep -Seconds 3

        #安裝新版
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\Vghtc\00_mis\$($new_pcsc_path.name)\gCIE_Setup\gCIE_Setup.msi /quiet /norestart" -Wait 

        #記錄內容
        $log_string = "updated,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
        $log_string | Add-Content -PassThru $log_file
        Send-LineNotifyMessage -Message $log_string
        Start-Sleep -Seconds 2
    }
    elseif ( ($installedPCSC -eq $null) -and `
        (test-path C:\NHI\sam\COMX1\0640140012001000005984.SAM) ) {
        #這是一個曾經發生但原因不明的bug, 雲端模組被不明原因移除了. 
        #所以安裝列表中不存在, 但SAM卻存在.
        #發現這情形就再裝一次

        #復制新版
        $new_pcsc_path = "\\172.20.5.187\mis\23-讀卡機控制軟體\CMS_CS5.1.5.7_20220925\CS5.1.5.7版_20220925"
    
        $new_pcsc_path = Get-Item $new_pcsc_path
        Copy-Item -Path $new_pcsc_path -Destination "C:\Vghtc\00_mis\$($new_pcsc_path.name)" -Recurse -Force -Verbose
       
        #安裝新版
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\Vghtc\00_mis\$($new_pcsc_path.name)\gCIE_Setup\gCIE_Setup.msi /quiet /norestart" -Wait 

        #記錄內容
        $log_string = "abnormal uninstall,PCSC:$($installedPCSC.version),$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
        $log_string | Add-Content -PassThru $log_file
        Send-LineNotifyMessage -Message $log_string
        Start-Sleep -Seconds 2
    }
    else {
        #不符升級條件
        
        if ($installedPCSC.Version -eq "5.1.57") { write-output "PCSC版本己是5.1.57." }
        if ($ipv4 -eq $null) { write-output "IP為非灣橋院區IP." }
    }

    #檢查SAM檔
    #升級應該不用檢查這個.

    #復制Link
    $diff1 = "C:\Users\Public\Desktop\雲端安全模組主控台.lnk"
    $diff2 = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\雲端安全模組主控台.lnk"
    $compare_result = Compare-Object -ReferenceObject (Get-Content $diff1) -DifferenceObject (Get-Content $diff2)

    if (($compare_result -ne $null) -and (Test-Path -Path $diff2)) {
        #如果原本就不在startup資料夾裡, 也不用復制過去.
        Copy-Item -path $diff1 -Destination $diff2 -Force
        $log_string = "link copied,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
        $log_string | Add-Content -PassThru $log_file
        Send-LineNotifyMessage -Message $log_string
        Start-Sleep -Seconds 2
    }

    #復制CSHIS.dll到指定資料夾

    $source_dll = Get-ItemProperty -Path "C:\NHI\LIB\CSHIS.dll" -ErrorAction SilentlyContinue

    if ($source_dll -ne $null) {
        
        $setup_file_ = @(
            "C:\VGHTC\ICCard\CSHIS.dll",
            "C:\ICCARD_HIS\CSHIS.dll",
            "C:\vhgp\ICCard\CSHIS.dll"
        )

        #關閉下列程式,以防占用DLL.
        Stop-Process -Name IccPrj -ErrorAction SilentlyContinue
        Stop-Process -Name HISLogin -ErrorAction SilentlyContinue
        Stop-Process -Name csfsim -ErrorAction SilentlyContinue
        
        $count = 0

        foreach ($i in $setup_file_) {

            $i_version = Get-ItemProperty -Path $i

            $result = $source_dll.VersionInfo.ProductVersion -ne $i_version.VersionInfo.ProductVersion
                
            if ($result) {
                Copy-Item -Path $source_dll.FullName -Destination $i -Force
                $count += 1
                $log_string = "$($source_dll.FullName):$($source_dll.VersionInfo.ProductVersion),>>,$i :$($i_version.VersionInfo.ProductVersion)"  
                $log_string | Add-Content -PassThru $log_file

            }
                    
        }
        if ($count -ne 0) {

            $log_string = "Dll copied$count,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file
            Send-LineNotifyMessage -Message $log_string
            Start-Sleep -Seconds 2
        }
    }       
    #copy dlls
    #就是復制"C:\NHI\LIB\"裡所有dll到3個資料夾.
    $setup_file_ = Get-ChildItem -Path "C:\NHI\LIB2\" -ErrorAction SilentlyContinue

    if ($setup_file_ -ne $null) {
    
        $setup_file_target_path = @(
            "C:\ICCARD_HIS",
            "C:\Windows\System32",
            "C:\Windows\System"    
        )

        $count = 0

        foreach ($i in $setup_file_) {
                
            foreach ($j in $setup_file_target_path) {
                $j_version = Get-ItemProperty -path "$j\$($i.name)" -ErrorAction SilentlyContinue
                $result = $i.VersionInfo.ProductVersion -ne $j_version.VersionInfo.ProductVersion
                if ($result) {
                    copy-item -Path $i.FullName -Destination $($j + "\" + $i.Name) -Force
                    $count += 1
                    $log_string = "$($i.FullName): $($i.VersionInfo.ProductVersion ),>>,$($j + "\" + $i.Name): $($j_version.VersionInfo.ProductVersion)"  
                    $log_string | Add-Content -PassThru $log_file
                }
            }
        }
        if ($count -ne 0) {
            $log_string = "LibDll copied$count,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file
            Send-LineNotifyMessage -Message $log_string
            Start-Sleep -Seconds 2
        }
    }

    #檢查系統環境變數
    $setting_file = "C:\VGHTC\00_mis\中榮iccard環境變數設定.bat"
    Write-Output "執行環境設定: $setting_file"
    $path = "C:\VGHTC\ICCard"
    $result = Check-EnvPathContains "C:\VGHTC\ICCard"

    if ($result -eq $false) {
    
        Write-Warning "系統環境變數中不包含 $path"
        
        if ($check_admin) {
            if (Test-Path $setting_file) {
                Write-Output "執行設定檔: $setting_file"
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c $setting_file"  -Wait
            }
        }
            
    }
    elseif ($result -eq $true) {

        Write-Output "系統環境變數中包含$path"
        Write-Output "不需執行 $setting_file "

    }

}    


function update-virtualhc {
    #升級虛擬健保卡SDK 2.5.4

    #只有有裝過的才需要升級
    #新的不用安裝, 只要復制資料夾再執行程式即可

    #取得舊版軟體
    $installed_vhc = get-installedprogramlist | Where-Object -FilterScript { $_.Displayname -like "虛擬健保卡*" }

    $log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VirtualHC_254.log"

    if ($installed_vhc -ne $null) {
            
        #復制2.5.4版到本機
        $vhc_path = "\\172.20.5.187\mis\25-虛擬健保卡\診間\VHIC_virtual-nhicard+SDK+Setup-2.5.4"
        $vhc_path = Get-Item $vhc_path
        Copy-Item -Path $vhc_path -Destination "c:\NHI\$($vhc_path.name)" -Recurse -Force -Verbose

        #移除軟體
        $unistalll_strign = $installed_vhc.QuietUninstallString.Split("""")
        Start-Process -FilePath $unistalll_strign[1] -ArgumentList $unistalll_strign[2] -Wait -ErrorAction SilentlyContinue -NoNewWindow
        #$installed_vhc.uninstall()
        Start-Sleep -Seconds 3

        #復制捷徑到桌面及啟動
        Create-Shortcut -TargetPath "C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\虛擬健保卡控制軟體-正式版.2.5.4.exe" -ShortcutPath "C:\users\public\desktop\虛擬健保卡控制軟體.lnk"
        Create-Shortcut -TargetPath "C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\虛擬健保卡控制軟體-正式版.2.5.4.exe" -ShortcutPath "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\虛擬健保卡控制軟體.lnk"
    

        #open firewall
        #netsh advfirewall firewall add rule name='Allow 虛擬健保卡控制軟體' dir=in action=allow program='C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\虛擬健保卡控制軟體-正式版.2.5.4.exe'
        Start-Process netsh.exe -ArgumentList "advfirewall firewall add rule name='Allow 虛擬健保卡控制軟體' dir=in action=allow program='C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\虛擬健保卡控制軟體-正式版.2.5.4.exe'"

        $log_string = "update virtualHC$count,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
        $log_string | Add-Content -PassThru $log_file
           
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

    if ($check_admin) { 
        update-pcsc
        update-virtualhc
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }

    #重新啟健保卡程式.
    Stop-Process -Name csfsim -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Start-Process -FilePath C:\nhi\BIN\csfsim.exe -ErrorAction SilentlyContinue

    Write-Output "升級完成."
    Start-Sleep -Seconds 20
}