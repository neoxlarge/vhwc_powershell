#升級虛擬健保卡SDK 2.5.4

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


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


function check-OPDList {
    # 檢查目前電腦名稱是否在升級名單(opd_list.json)內. 有的話,回傳資料.
   
    # 將變數 $json_file 設定為 JSON 檔案 "opd_list.json" 的路徑
    $path = get-item -Path $PSCommandPath
    $json_file = "$($Path.DirectoryName)\opd_list.json"
    #Write-Output $json_file
    
    # 電腦名稱
    $pc_name = $env:COMPUTERNAME

    # 讀取 JSON 檔案的內容並將其儲存在變數 $json_content 中
    $json_content = Get-Content -Path $json_file -Raw

    # 將 JSON 內容轉換為 PowerShell 物件並將其指派給變數 $opd_json
    $opd_json = ConvertFrom-Json -InputObject $json_content

    # 初始化變數 $opd，並將其設定為 null
    $opd = $null

    # 找出符合電腦名稱的資料.
    foreach ($o in $opd_json.psobject.properties) {

        $result = $o.Value.name -eq $pc_name
   
        if ($result) {
            Write-Output $o.Value
            return $o.Value
            break
        }
 
    }
}

function install-virtualnhc {
    #升級虛擬健保卡SDK 2.5.4

    #升級條件:
    # 1. 診間醫生電腦才需要安裝.
    # 2. 診間醫生電腦目前皆為win10.
    # 所以要參考安裝列表opd_list.json 
    
    # 新版的虛擬健保卡SDK 2.5.4不用安裝, 只要復制資料夾再執行程式即可

    # 檢查目前電腦名稱是否在升級名單(opd_list.json)內.
    # 底下check-OPDList為powershell v5 語法,無法在win7中執行. 就先限定在win10執行.
    
    if ($(Get-OSVersion) -eq "Windows 10") {

        $ipv4 = Get-IPv4Address 

        $opd = check-OPDList

        #if ($opd.virturl_NHIcard) {
        if ($true) {    
            #如果virtual_NHIcar值是$true 表示可升級.
            #這版2.5.4, 不好判斷是否己經安裝, 就直接再裝一次.
            #診間醫師電腦都有最高權限, 在此就暫不考慮權限問題. 一般使用者會無法使用.

            #取得舊版軟體
            $installed_vhc = get-installedprogramlist | Where-Object -FilterScript { $_.Displayname -like "虛擬健保卡*" }

            $log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VirtualHC_254.log"
            
            
            #移除舊版軟體
            if ($installed_vhc -ne $null) {
                
                $unistalll_strign = $installed_vhc.QuietUninstallString.Split("""")
                Start-Process -FilePath $unistalll_strign[1] -ArgumentList $unistalll_strign[2] -Wait -ErrorAction SilentlyContinue -NoNewWindow
                #$installed_vhc.uninstall()
                Start-Sleep -Seconds 3
            }


            #檢查新版是否己經復制2.5.4版到本機
            $vhc_path = "\\172.20.5.187\mis\25-虛擬健保卡\診間\VHIC_virtual-nhicard+SDK+Setup-2.5.4"
            $vhc_path = Get-Item $vhc_path

            Copy-Item -Path $vhc_path -Destination "c:\NHI\$($vhc_path.name)" -Recurse -Force -Verbose
                            
            #復制捷徑到桌面及啟動
            $target_path = "C:\users\public\desktop\虛擬健保卡控制軟體.lnk"
            $result = !(Test-Path -Path $target_path) -and $check_admin
            if ($result) {
            Create-Shortcut -TargetPath "C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\虛擬健保卡控制軟體-正式版.2.5.4.exe" -ShortcutPath $target_path
            #暫時拿掉啟動的, 不是所有電腦有需要跑
            #Create-Shortcut -TargetPath "C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\虛擬健保卡控制軟體-正式版.2.5.4.exe" -ShortcutPath "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\虛擬健保卡控制軟體.lnk"
            }

            #open firewall
            #netsh advfirewall firewall add rule name='Allow 虛擬健保卡控制軟體' dir=in action=allow program='C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\虛擬健保卡控制軟體-正式版.2.5.4.exe'
            $exe_path1 = "$($env:USERPROFILE)\AppData\Local\Temp\C\vNHI\虛擬健保卡控制軟體.exe"
            $exe_path2 = "$($env:USERPROFILE)\AppData\Local\Temp\C\vNHI\resources\app\VHCNHI_Slient\VHCNHI_Slient.exe"

            $rule_name1 = "Allow 虛擬健保卡控制軟體 ($($env:USERNAME) main) "
            $rule_name2 = "Allow 虛擬健保卡控制軟體 ($($env:USERNAME) VHCNHI_Silent)"

            #管理者權限vhwcmis的證書.
            $Username = "vhwcmis"
            $Password = "Mis20190610"
            $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
            
            <#
            $cimsession = New-CimSession -ComputerName $env:COMPUTERNAME -Credential $credential

            if (Get-NetFirewallRule -DisplayName $rule_name1 -CimSession $cimsession -ErrorAction SilentlyContinue) {
                Remove-NetFirewallRule -DisplayName -CimSession $cimsession $rule_name1
            }
            if (Get-NetFirewallRule -DisplayName $rule_name2 -CimSession $cimsession -ErrorAction SilentlyContinue) {
                Remove-NetFirewallRule -DisplayName $rule_name2 -CimSession $cimsession
            }

            $arg1 = "advfirewall firewall add rule name=""$rule_name1 "" dir=in action=allow program=""$exe_path1"""
            Write-Output  "建立firewall rule: $arg1"
            #Start-Process netsh.exe -ArgumentList $arg1 -NoNewWindow -wait -Credential $credential
            Invoke-Command -ComputerName $env:COMPUTERNAME -Credential $credential -ScriptBlock {
                param ($argument1)
                netsh.exe $argument1
            } -ArgumentList $arg1



            $arg2 = "advfirewall firewall add rule name=""$rule_name2"" dir=in action=allow program=""$exe_path2"""
            write-output "建立firewall rule: $arg2"
            #Start-Process netsh.exe -ArgumentList $arg2 -NoNewWindow -wait -Credential $credential
            Invoke-Command -ComputerName $env:COMPUTERNAME -Credential $credential -ScriptBlock {
                param ($argument2)
                netsh.exe $argument2
            } -ArgumentList $arg2

#>

            $log_string = "update V-NHICard,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file

        }
    
            
        
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
        install-virtualnhc
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}