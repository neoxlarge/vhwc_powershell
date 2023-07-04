# 升級雲端安控元件健保卡讀卡機控制(PCSC 5.1.5.7)

param($runadmin)


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
    $installedVersions = "5.1.53", "5.1.55"
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
    
        $log_string = "Boot,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)"
        $log_string | Add-Content -PassThru $log_file
    
        #符合升級條件,進行升級
        if ($check_condition) {
                
            #移除舊版
            $installedPCSC.uninstall()

            #復制新版
            $new_pcsc_path = "\\172.20.5.187\mis\23-讀卡機控制軟體\CMS_CS5.1.5.7_20220925\CS5.1.5.7版_20220925"
    
            $new_pcsc_path = Get-Item $new_pcsc_path
            Copy-Item -Path $new_pcsc_path -Destination "C:\Vghtc\00_mis\$($new_pcsc_path.name)" -Recurse -Force -Verbose

            Stop-Process -Name csfsim -Force -ErrorAction SilentlyContinue

            #安裝新版
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\Vghtc\00_mis\$($new_pcsc_path.name)\gCIE_Setup\gCIE_Setup.msi /quiet /norestart" -Wait 

            #重新啟健保卡程式.
            Stop-Process -Name csfsim -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
            Start-Process -FilePath C:\nhi\BIN\csfsim.exe -ErrorAction SilentlyContinue

            #記錄內容
            $log_string = "updated,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file
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

        if ($compare_result -ne $null) {
            Copy-Item -path $diff1 -Destination $diff2 -Force
            $log_string = "link copied,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file
        }

        #復制CSHIS.dll到指定資料夾

        $source_dll = Get-ItemProperty -Path "C:\NHI\LIB\CSHIS.dll"
        
        $setup_file_ = @(
            "C:\VGHTC\ICCard\CsHis.dll",
            "C:\ICCARD_HIS\CsHis.dll",
            "C:\vhgp\ICCard\CsHis.dll"
        )

        $count = 0

        foreach ($i in $setup_file_) {

            $i_version = Get-ItemProperty -Path $i

            $result = $source_dll.VersionInfo.ProductVersion -ne $i_version.VersionInfo.ProductVersion
            
            if ($result) {
                Copy-Item -Path $source_dll.FullName -Destination $i -Force
                $count += 1
            }
                
        }
        if ($count -ne 0) {
            $log_string = "Dll copied$count,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file
        }
        
        #copy dlls
        #就是復制"C:\NHI\LIB\"裡所有dll到3個資料夾.
        $setup_file_ = Get-ChildItem -Path "C:\NHI\LIB\"
    
        $setup_file_target_path = @(
            "C:\ICCARD_HIS",
            "C:\Windows\System32",
            "C:\Windows\System"    
        )

        $count = 0

        foreach ($i in $setup_file_) {
            
            foreach ($j in $setup_file_target_path) {
                $result = $i.VersionInfo.ProductVersion -ne $(Get-ItemProperty -path ("$j\$($i.name)")).VersionInfo.ProductVersion
                if ($result) {
                    copy-item -Path $i.FullName -Destination ($j + "\" + $i.Name) -Force
                    $count += 1
                }
            }
        }
        if ($count -ne 0) {
            $log_string = "LibDll copied$count,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
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
        }
        else {
            Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
        }
        Write-Output "升級完成."
        Start-Sleep -Seconds 20
    }