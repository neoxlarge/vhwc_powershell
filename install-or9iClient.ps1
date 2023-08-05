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

    $responsefile = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\03-Oracle_ora\silentinstall.rsp"
    $ImagePath = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\01-oracle9i_client.iso"

    if ($check_or9i -eq $false) {
        Write-Host "Oracle 9i Client �i��w��:"

        #����oracle 9i client iso ��. �A��ϵw���N�����x:
        $MountedDisk = Mount-DiskImage -ImagePath $ImagePath -PassThru
        $DriverLetter = ($MountedDisk | Get-Volume).DriveLetter
        # ����ݭn���N��������
        $partition = Get-Partition -DriveLetter "$($DriverLetter):"
        # �����즳���Ϻо��N��
        Remove-PartitionAccessPath -Partition $partition -AccessPath "$($DriverLetter):"
        # ���t�s���Ϻо��N��
        Set-Partition -Partition $partition -NewDriveLetter x:


        if ($Error) {
            # ����x�W��X���~����
            #�`�N�A�b�Y�Ǳ��p�U�A�p�G���������\�A�h Mount-DiskImage �R�O�i�ण�|��^����ȡC�b�o�ر��p�U�A�z�i�H�q�L�ˬd $Error �ܶq�ӽT�w�O�_�o�Ϳ��~
            $Error[0].Exception.Message
        }

        #�_��responsefile �쥻��, �ç�s���|�쥻��
        Copy-Item -Path $responsefile -Destination $env:TEMP -Force -Verbose
        $responsefile = "$env:TEMP\$($responsefile.Split("\")[-1])"

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
            Write-Output "wait log"
        } until ( $log -ne $null )

        do {
            $proc = Get-Process -Name javaw -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 100
        } until ( !$proc )
        #silent�w�˵���
    
        #�ק������ܼ�
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

        Copy-Item -Path $ora -Destination "C:\Oracle\network\ADMIN\" -Force -Verbose

        Write-Output "Oracle 9i Client �w�˵���."

    }

}


function install-BDE {
    #install Borland DataBase Engine


    $software_name = "Borland DataBase Engine"
    $software_path = "\\172.20.1.122\share\software\00newpc\01-Oracle9i_BDE\02-BDE DISK.zip"

    ## ��X�n��O�_�v�w��
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_name -eq $null) {

        Write-Host "�w��Borland DataBase Engine:"
        #unzip file
        Expand-Archive -Path $zip_file -DestinationPath "$($env:temp)\BDE"
        $install_exe = Get-ChildItem -Path "$($env:temp)\BDE" -Name setup.exe -Recurse

        Start-Process -FilePath "$($env:temp)\BDE\$install_exe" -ArgumentList "/S /V/passive" -Wait
        Start-Sleep -Seconds 2


        #�_��]�w��
        #�]�w�ɦ���win7, win10
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

        #�վ�]�w��
        $BDE_path = "HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\settigs\DRIVERS\ORACLE\INIT\"

        Set-ItemProperty -Path $BDE_path -Name "DLL32" -Value "SQLORA8.DLL"
        Set-ItemProperty -Path $BDE_path -Name "VENDOR" -Value "OCI.DLL"

        Write-Output "Borland DataBase Engine �w�˵���."

    }
   
    

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
        install-or9iclient    
        install-BDE
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}