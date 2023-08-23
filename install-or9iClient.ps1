# �w��Oracle 9i Client

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
    #�w��Oracle 9i Client

    #�ˬdOracle 9i ��lient �O�_�v�g�w��
    #�ˬd�M��:
    #1. c:\oracle\ora92\bin�O�_�s�b
    #2. C:\Program Files\Oracle\jre�O�_�s�b, jre�|�s�POUI(oracle universal installer)�@�_�w��.

    #���U���i�ΨӧP�_Oracle�O�_�v�w��, ����orace 9i�ᤴ�|�d�U.
    #1. HKLM:\SOFTWARE\ORACLE �� HKLM:\SOFTWARE\WOW6432Node\ORACLE �O�_�s�b
    #2. c:\oracle\ora92 �O�_�s�b

    $check_or9i = Test-Path -Path "c:\oracle\ora92\bin"
    
    #responsefile ��Oracle univeral Install �۰ʦw�� silent install�ɥ���.
    $responsefile = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-Oracle_ora\silentinstall.rsp"
    #iso��
    $ImagePath = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\01-oracle9i_client.iso"

    if ($check_or9i -eq $false) {
        Write-Host "Oracle 9i Client �i��w��:"

        #����oracle 9i client iso ��. 
        $MountedDisk = Mount-DiskImage -ImagePath $ImagePath -PassThru
        $DriverLetter = ($MountedDisk | Get-Volume).DriveLetter

        if ($Error) {
            # ����x�W��X���~����
            #�`�N�A�b�Y�Ǳ��p�U�A�p�G���������\�A�h Mount-DiskImage �R�O�i�ण�|��^����ȡC�b�o�ر��p�U�A�z�i�H�q�L�ˬd $Error �ܶq�ӽT�w�O�_�o�Ϳ��~
            $Error[0].Exception.Message
        }


        #�_��responsefile �쥻��, �ç�s���|�쥻��
        Copy-Item -Path $responsefile -Destination $env:TEMP -Force
        $responsefile = "$env:TEMP\$($responsefile.Split("\")[-1])"

        #�w���ɪ����|�M�w�˰Ѽ�
        $install_exe = "$($DriverLetter):\install\win32\setup.exe"
        $install_arrg = "-silent -nowelcome -responseFile $responsefile"


        #�w�˷|�I�sjavaw.exe�Ӧw�ˡA�åB�g�Jlog.
        #���ˬdlog��, �p�G����ܦw�˵{���}��, 
        #���ˬdjavaw����,�Y�w�˵���.
    
        #�M��log
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "x86" { $log_path = "$($env:ProgramFiles)\Oracle\Inventory\logs" }
            "AMD64" { $log_path = "${env:ProgramFiles(x86)}\Oracle\Inventory\logs" }
        }

        Remove-Item "$log_path\." -Recurse -Force -ErrorAction SilentlyContinue

        #����w��oracle 9i client
        Start-Process -FilePath $install_exe -ArgumentList $install_arrg -Wait 

        do {
            $log = Get-ChildItem -Path $log_path -File "*.log" -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 100
            Write-Output "����Oracle 9i Client �w�ˤ�: wait log"
        } until ( $log -ne $null )

        do {
            $proc = Get-Process -Name javaw -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 100
            Write-Host "����Oracle 9i Client �w�ˤ�: wait javaw"
        } until ( $proc -eq $null)
        
        #silent�w�˵���

        #unmount iso
        Dismount-DiskImage -ImagePath $ImagePath
    
        #�ק�t�������ܼ�
        #�ˬd�M��: �������s�b, jre���|��, �|��net manager�}���_��.
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
    
        #�_��oracle 9i �]�w��
        $ora = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-Oracle_ora\tnsnames.ora"

        Copy-Item -Path $ora -Destination "C:\oracle\ora92\network\ADMIN" -Force -Verbose

        Write-Output "Oracle 9i Client �w�˵���."

    }
    else {
        Write-Host "Oracle 9i Client �v�w��."
    }

}


function install-BDE {
    #install Borland DataBase Engine


    $software_name = "Borland DataBase Engine"
    $software_path = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\02-BDE DISK.zip"

    ## ��X�n��O�_�v�w��
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {

        Write-Host "�w��Borland DataBase Engine:"
        #unzip file
        Expand-Archive -Path $software_path -DestinationPath "$($env:temp)\BDE" -force
        $install_exe = Get-ChildItem -Path "$($env:temp)\BDE" -Name setup.exe -Recurse

        #BDE���w�˦��G���Oborland��l���w��, �ҥH/S���hsilent install�@��, 
        #��msiexec ��/passive�w?, �A�ɤW�ʤ֪�����.
        
        #����w��
        Start-Process -FilePath "$($env:temp)\BDE\$install_exe" -ArgumentList "/S /V/passive" -Wait
        
        Start-Sleep -Seconds 2

        #�_��BDE��Ƨ�, �䤤��SQLORA8.DLL���ɮ�
        Copy-Item -Path "$($env:temp)\BDE\BDE DISK\Common\Borland Shared\BDE\*" -Destination "C:\Program Files (x86)\Common Files\Borland Shared\BDE\" -Force -Recurse   

        #�_��]�w��
        #�]�w�ɦ���win7, win10
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

        #20230823, add copy pdoxsusr.net �H�קK�X�{unknow table type ���~. �n�}��Ū���v�����ϥΪ�, �o�bgrant-fullcontrolpermission�@.
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "x86" { $dest_path = "$($env:ProgramFiles)\Common Files\Borland Shared\BDE\" }
            "AMD64" { $dest_path = "$(${env:ProgramFiles(x86)})\Common Files\Borland Shared\BDE\" }
            default { $dest_path = "$(${env:ProgramFiles(x86)})\Common Files\Borland Shared\BDE\" }
        }

        Copy-Item -Path "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-BDE_cfg\PDOXUSRS.NET" -Destination $dest_path -Force

        #�վ�]�w��, �۰ʦw�ˤ��|�إߩ��U��registry, �ۦ�إ�
        $registryPath1 = "HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\DRIVERS\ORACLE\DB OPEN"
        $registryPath2 = "HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\DRIVERS\ORACLE\INIT"
        $registryPath3 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\SharedDlls"
       
        # �ˬd�óЫظ��|
        if (!(Test-Path $registryPath1)) {
            New-Item -Path $registryPath1 -Force | Out-Null
        }
        if (!(Test-Path $registryPath2)) {
            New-Item -Path $registryPath2 -Force | Out-Null
        }
       
        # �]�m���U����ȹ�
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
      
        #20230818, wadm-reg-pc01 ������112�~�sPC�󴫤@�x, ��@��m�X�{�L�k�P�ɹ��2�ӦP�ɨϥ�BDE�����~BDE $210D, ���~�i�H�Ѧҩ��UURL�ѨM.
        # https://techjourney.net/error-2501-210d-while-attempting-to-initialize-borland-database-engine-bde/
        # https://www.fox-saying.com/blog/post/44256505
        #
       
        $registryPath4 = "HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT"
        Set-ItemProperty -Path $registryPath4 -Name "SHAREDMEMLOCATION" -Value "0x5BDE"
        Set-ItemProperty -Path $registryPath4 -Name "SHAREDMEMSIZE" -Value "4096"
        #���U�o�Ӷ���諸,���T�w�O�����v?
        Set-ItemProperty -Path $registryPath4 -Name "LANGDRIVER" -Value "taiwan"
        

        # Borland DataBase Engine �w�˵���

        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
        

    }
    
    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)
    

}


#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q���J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        #�S���޲z���v��������.
        install-or9iclient    
        install-BDE
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw��oracle9i client & BDE �n��, �ХH�޲z���b������."
    }
    pause
}