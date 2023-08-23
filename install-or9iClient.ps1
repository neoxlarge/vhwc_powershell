# 安裝Oracle 9i Client

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


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


function install-or9iclient {
    #安裝Oracle 9i Client

    #檢查Oracle 9i Ｃlient 是否己經安裝
    #檢查清單:
    #1. c:\oracle\ora92\bin是否存在
    #2. C:\Program Files\Oracle\jre是否存在, jre會連同OUI(oracle universal installer)一起安裝.

    #底下不可用來判斷Oracle是否己安裝, 移除orace 9i後仍會留下.
    #1. HKLM:\SOFTWARE\ORACLE 或 HKLM:\SOFTWARE\WOW6432Node\ORACLE 是否存在
    #2. c:\oracle\ora92 是否存在

    $check_or9i = Test-Path -Path "c:\oracle\ora92\bin"
    
    #responsefile 為Oracle univeral Install 自動安裝 silent install時必須.
    $responsefile = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-Oracle_ora\silentinstall.rsp"
    #iso檔
    $ImagePath = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\01-oracle9i_client.iso"

    if ($check_or9i -eq $false) {
        Write-Host "Oracle 9i Client 進行安裝:"

        #掛載oracle 9i client iso 檔. 
        $MountedDisk = Mount-DiskImage -ImagePath $ImagePath -PassThru
        $DriverLetter = ($MountedDisk | Get-Volume).DriveLetter

        if ($Error) {
            # 控制台上輸出錯誤消息
            #注意，在某些情況下，如果掛載未成功，則 Mount-DiskImage 命令可能不會返回任何值。在這種情況下，您可以通過檢查 $Error 變量來確定是否發生錯誤
            $Error[0].Exception.Message
        }


        #復制responsefile 到本機, 並更新路徑到本機
        Copy-Item -Path $responsefile -Destination $env:TEMP -Force
        $responsefile = "$env:TEMP\$($responsefile.Split("\")[-1])"

        #安裝檔的路徑和安裝參數
        $install_exe = "$($DriverLetter):\install\win32\setup.exe"
        $install_arrg = "-silent -nowelcome -responseFile $responsefile"


        #安裝會呼叫javaw.exe來安裝，並且寫入log.
        #先檢查log檔, 如果有表示安裝程式開啟, 
        #當檢查javaw結束,即安裝結束.
    
        #清空log
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "x86" { $log_path = "$($env:ProgramFiles)\Oracle\Inventory\logs" }
            "AMD64" { $log_path = "${env:ProgramFiles(x86)}\Oracle\Inventory\logs" }
        }

        Remove-Item "$log_path\." -Recurse -Force -ErrorAction SilentlyContinue

        #執行安裝oracle 9i client
        Start-Process -FilePath $install_exe -ArgumentList $install_arrg -Wait 

        do {
            $log = Get-ChildItem -Path $log_path -File "*.log" -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 100
            Write-Output "等待Oracle 9i Client 安裝中: wait log"
        } until ( $log -ne $null )

        do {
            $proc = Get-Process -Name javaw -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 100
            Write-Host "等待Oracle 9i Client 安裝中: wait javaw"
        } until ( $proc -eq $null)
        
        #silent安裝結束

        #unmount iso
        Dismount-DiskImage -ImagePath $ImagePath
    
        #修改系統環境變數
        #檢查清單: 必須都存在, jre路徑錯, 會讓net manager開不起來.
        #C:\oracle\ora92\bin
        #C:\Program Files\Oracle\jre\1.3.1\bin
        #C:\Program Files\Oracle\jre\1.1.8\bin
        #C:\Program Files\Oracle\oui\bin

        $pathsToCheck = @(
            "C:\oracle\ora92\bin",
            "C:\Program Files\Oracle\jre\1.3.1\bin",
            "C:\Program Files\Oracle\jre\1.1.8\bin",
            "C:\Program Files\Oracle\oui\bin"
        )
    
        $currentPaths = $env:Path -split ';'
    
        foreach ($path in $pathsToCheck) {
            if (-not $currentPaths.Contains($path)) {
                $currentPaths += $path
                #Write-Host "Added path: $path"
            }
            else {
                #Write-Host "Path already exists: $path"
            }
        }
    
        $newPath = $currentPaths -join ';'
        [Environment]::SetEnvironmentVariable("Path", $newPath, [EnvironmentVariableTarget]::Machine)
    
        #復制oracle 9i 設定檔
        $ora = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-Oracle_ora\tnsnames.ora"

        Copy-Item -Path $ora -Destination "C:\oracle\ora92\network\ADMIN" -Force -Verbose

        Write-Output "Oracle 9i Client 安裝結束."

    }
    else {
        Write-Host "Oracle 9i Client 己安裝."
    }

}


function install-BDE {
    #install Borland DataBase Engine


    $software_name = "Borland DataBase Engine"
    $software_path = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\02-BDE DISK.zip"

    ## 找出軟體是否己安裝
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {

        Write-Host "安裝Borland DataBase Engine:"
        #unzip file
        Expand-Archive -Path $software_path -DestinationPath "$($env:temp)\BDE" -force
        $install_exe = Get-ChildItem -Path "$($env:temp)\BDE" -Name setup.exe -Recurse

        #BDE的安裝似乎不是borland原始的安裝, 所以/S失去silent install作用, 
        #用msiexec 的/passive安?, 再補上缺少的部分.
        
        #執行安裝
        Start-Process -FilePath "$($env:temp)\BDE\$install_exe" -ArgumentList "/S /V/passive" -Wait
        
        Start-Sleep -Seconds 2

        #復制BDE資料夾, 其中有SQLORA8.DLL等檔案
        Copy-Item -Path "$($env:temp)\BDE\BDE DISK\Common\Borland Shared\BDE\*" -Destination "C:\Program Files (x86)\Common Files\Borland Shared\BDE\" -Force -Recurse   

        #復制設定檔
        #設定檔有分win7, win10
        $os = Get-OSVersion

        switch ($os) {
            "Windows 7" { $cfg_file = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-BDE_cfg\idapi32-Win7-1100203.cfg" }
            "Windows 10" { $cfg_file = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-BDE_cfg\idapi32-Win10-1100217.cfg" } 
            default { $cfg_file = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-BDE_cfg\idapi32-Win10-1100217.cfg" }
        }

        switch ($env:PROCESSOR_ARCHITECTURE) {
            "x86" { $dest_path = "$($env:ProgramFiles)\Common Files\Borland Shared\BDE\idapi32.cfg" }
            "AMD64" { $dest_path = "$(${env:ProgramFiles(x86)})\Common Files\Borland Shared\BDE\idapi32.cfg" }
            default { $dest_path = "$(${env:ProgramFiles(x86)})\Common Files\Borland Shared\BDE\idapi32.cfg" }
        }
        
        Copy-Item -Path $cfg_file -Destination $dest_path -Force

        #20230823, add copy pdoxsusr.net 以避免出現unknow table type 錯誤. 要開啟讀取權限給使用者, 這在grant-fullcontrolpermission作.
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "x86" { $dest_path = "$($env:ProgramFiles)\Common Files\Borland Shared\BDE\" }
            "AMD64" { $dest_path = "$(${env:ProgramFiles(x86)})\Common Files\Borland Shared\BDE\" }
            default { $dest_path = "$(${env:ProgramFiles(x86)})\Common Files\Borland Shared\BDE\" }
        }

        Copy-Item -Path "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-BDE_cfg\PDOXUSRS.NET" -Destination $dest_path -Force

        #調整設定值, 自動安裝不會建立底下的registry, 自行建立
        $registryPath1 = "HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\DRIVERS\ORACLE\DB OPEN"
        $registryPath2 = "HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\DRIVERS\ORACLE\INIT"
        $registryPath3 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\SharedDlls"
       
        # 檢查並創建路徑
        if (!(Test-Path $registryPath1)) {
            New-Item -Path $registryPath1 -Force | Out-Null
        }
        if (!(Test-Path $registryPath2)) {
            New-Item -Path $registryPath2 -Force | Out-Null
        }
       
        # 設置註冊表鍵值對
        Set-ItemProperty -Path $registryPath1 -Name "SERVER NAME" -Value "ORA_SERVER"
        Set-ItemProperty -Path $registryPath1 -Name "USER NAME" -Value "MYNAME"
        Set-ItemProperty -Path $registryPath1 -Name "NET PROTOCOL" -Value "TNS"
        Set-ItemProperty -Path $registryPath1 -Name "OPEN MODE" -Value "READ/WRITE"
        Set-ItemProperty -Path $registryPath1 -Name "SCHEMA CACHE SIZE" -Value "8"
        Set-ItemProperty -Path $registryPath1 -Name "LANGDRIVER" -Value ""
        Set-ItemProperty -Path $registryPath1 -Name "SQLQRYMODE" -Value ""
        Set-ItemProperty -Path $registryPath1 -Name "SQLPASSTHRU MODE" -Value "SHARED AUTOCOMMIT"
        Set-ItemProperty -Path $registryPath1 -Name "SCHEMA CACHE TIME" -Value "-1"
        Set-ItemProperty -Path $registryPath1 -Name "MAX ROWS" -Value "-1"
        Set-ItemProperty -Path $registryPath1 -Name "BATCH COUNT" -Value "200"
        Set-ItemProperty -Path $registryPath1 -Name "ENABLE SCHEMA CACHE" -Value "FALSE"
        Set-ItemProperty -Path $registryPath1 -Name "SCHEMA CACHE DIR" -Value ""
        Set-ItemProperty -Path $registryPath1 -Name "ENABLE BCD" -Value "FALSE"
        Set-ItemProperty -Path $registryPath1 -Name "ENABLE INTEGERS" -Value "FALSE"
        Set-ItemProperty -Path $registryPath1 -Name "LIST SYNONYMS" -Value "NONE"
        Set-ItemProperty -Path $registryPath1 -Name "ROWSET SIZE" -Value "20"
        Set-ItemProperty -Path $registryPath1 -Name "BLOBS TO CACHE" -Value "64"
        Set-ItemProperty -Path $registryPath1 -Name "BLOB SIZE" -Value "32"
        Set-ItemProperty -Path $registryPath1 -Name "OBJECT MODE" -Value "TRUE"
       
        Set-ItemProperty -Path $registryPath2 -Name "VERSION" -Value "4.0"
        Set-ItemProperty -Path $registryPath2 -Name "TYPE" -Value "SERVER"
        Set-ItemProperty -Path $registryPath2 -Name "DLL32" -Value "SQLORA8.DLL"
        Set-ItemProperty -Path $registryPath2 -Name "VENDOR INIT" -Value "OCI.DLL"
        Set-ItemProperty -Path $registryPath2 -Name "DRIVER FLAGS" -Value ""
        Set-ItemProperty -Path $registryPath2 -Name "TRACE MODE" -Value "0"
        Set-ItemProperty -Path $registryPath2 -Name "VENDOR" -Value "OCI.DLL"

        Set-ItemProperty -Path $registryPath3 -Name "C:\\Program Files (x86)\\Common Files\\Borland Shared\\BDE\\sqlora32.dll" -Value "00000001" -Type DWORD
        Set-ItemProperty -Path $registryPath3 -Name "C:\\Program Files (x86)\\Common Files\\Borland Shared\\BDE\\sqlora8.dll" -Value "00000001" -Type DWORD
        Set-ItemProperty -Path $registryPath3 -Name "C:\\Program Files (x86)\\Common Files\\Borland Shared\\BDE\\BDEADMIN.TOC" -Value "00000001" -Type DWORD
        Set-ItemProperty -Path $registryPath3 -Name "C:\\Program Files (x86)\\Common Files\\Borland Shared\\BDE\\sql_ora.cnf" -Value "00000001" -Type DWORD
        Set-ItemProperty -Path $registryPath3 -Name "C:\\Program Files (x86)\\Common Files\\Borland Shared\\BDE\\sql_ora8.cnf" -Value "00000001" -Type DWORD
      
        #20230818, wadm-reg-pc01 掛號室112年新PC更換一台, 曉婷位置出現無法同時實執2個同時使用BDE的錯誤BDE $210D, 錯誤可以參考底下URL解決.
        # https://techjourney.net/error-2501-210d-while-attempting-to-initialize-borland-database-engine-bde/
        # https://www.fox-saying.com/blog/post/44256505
        #
       
        $registryPath4 = "HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT"
        Set-ItemProperty -Path $registryPath4 -Name "SHAREDMEMLOCATION" -Value "0x5BDE"
        Set-ItemProperty -Path $registryPath4 -Name "SHAREDMEMSIZE" -Value "4096"
        #底下這個順手改的,不確定是不有影?
        Set-ItemProperty -Path $registryPath4 -Name "LANGDRIVER" -Value "taiwan"
        

        # Borland DataBase Engine 安裝結束

        #安裝完, 再重新取得安裝資訊
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
        

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
        #沒有管理員權限不執行.
        install-or9iclient    
        install-BDE
    }
    else {
        Write-Warning "無法取得管理員權限來安裝oracle9i client & BDE 軟體, 請以管理員帳號重試."
    }
    pause
}