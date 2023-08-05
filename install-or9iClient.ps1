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

    $responsefile = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-Oracle_ora\silentinstall.rsp"
    $ImagePath = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\01-oracle9i_client.iso"

    if ($check_or9i -eq $false) {
        Write-Host "Oracle 9i Client 進行安裝:"

        #掛載oracle 9i client iso 檔. 再把磁硬機代號改到x:
        $MountedDisk = Mount-DiskImage -ImagePath $ImagePath -PassThru
        $DriverLetter = ($MountedDisk | Get-Volume).DriveLetter
        # 獲取需要更改代號的分區
        $partition = Get-Partition -DriveLetter "$($DriverLetter):"
        # 移除原有的磁碟機代號
        Remove-PartitionAccessPath -Partition $partition -AccessPath "$($DriverLetter):"
        # 分配新的磁碟機代號
        Set-Partition -Partition $partition -NewDriveLetter x:


        if ($Error) {
            # 控制台上輸出錯誤消息
            #注意，在某些情況下，如果掛載未成功，則 Mount-DiskImage 命令可能不會返回任何值。在這種情況下，您可以通過檢查 $Error 變量來確定是否發生錯誤
            $Error[0].Exception.Message
        }

        #復制responsefile 到本機, 並更新路徑到本機
        Copy-Item -Path $responsefile -Destination $env:TEMP -Force -Verbose
        $responsefile = "$env:TEMP\$($responsefile.Split("\")[-1])"

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
            Write-Output "wait log"
        } until ( $log -ne $null )

        do {
            $proc = Get-Process -Name javaw -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 100
        } until ( !$proc )
        #silent安裝結束
    
        #修改環境變數
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
    
        #復制oracle 9i 設定當
        $ora = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-Oracle_ora\tnsnames.ora"

        Copy-Item -Path $ora -Destination "C:\Oracle\network\ADMIN\" -Force -Verbose

        Write-Output "Oracle 9i Client 安裝結束."

    }

}


function install-BDE {
    #install Borland DataBase Engine


    $software_name = "Borland DataBase Engine"
    $software_path = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\02-BDE DISK.zip"

    ## 找出軟體是否己安裝
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_name -eq $null) {

        Write-Host "安裝Borland DataBase Engine:"
        #unzip file
        Expand-Archive -Path $zip_file -DestinationPath "$($env:temp)\BDE"
        $install_exe = Get-ChildItem -Path "$($env:temp)\BDE" -Name setup.exe -Recurse

        Start-Process -FilePath "$($env:temp)\BDE\$install_exe" -ArgumentList "/S /V/passive" -Wait
        Start-Sleep -Seconds 2


        #復制設定檔
        #設定檔有分win7, win10
        $os = Get-OSVersion

        switch ($os) {
            "Windows 7" { $cfg_file = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-BDE_cfg\idapi32-Win7-1100203.cfg" }
            "Windows 10" { $cfg_file = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-BDE_cfg\idapi32-Win10-1100217.cfg" } 
            default { $cfg_file = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-BDE_cfg\idapi32-Win10-1100217.cfg" }
        }

        switch ($env:PROCESSOR_ARCHITECTURE) {
            "x86" { $dest_path = "$($env:ProgramFiles)\Common Files\Borland Shred\BDE\idapi32.cfg" }
            "AMD64" { $dest_path = "$(${env:ProgramFiles(x86)})\Common Files\Borland Shred\BDE\idapi32.cfg" }
            default { $dest_path = "$(${env:ProgramFiles(x86)})\Common Files\Borland Shred\BDE\idapi32.cfg" }

        }
        
        Copy-Item -Path $cfg_file -Destination $dest_path -Force

        #調整設定值
        $BDE_path = "HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\settigs\DRIVERS\ORACLE\INIT\"

        Set-ItemProperty -Path $BDE_path -Name "DLL32" -Value "SQLORA8.DLL"
        Set-ItemProperty -Path $BDE_path -Name "VENDOR" -Value "OCI.DLL"

        Write-Output "Borland DataBase Engine 安裝結束."

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

    if ($check_admin) { 
        install-or9iclient    
        install-BDE
    }
    else {
        Write-Warning "無法取得管理員權限來安裝軟體, 請以管理員帳號重試."
    }
    pause
}