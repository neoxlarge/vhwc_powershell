# �ɯŶ��ݦw�����󰷫O�dŪ�d������(PCSC 5.1.5.7)

param($runadmin)


function Get-IPv4Address {
    <#
    �^�ǧ�쪺IP,�u��b172.*�~���. 
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
    Write-Host "�ɯŰ��O�dŪ�d������n��PCSC 5.1.5.7"
    # �n�������ˬd
    # 1. ���o�v�w�˪���
    $installedPCSC = Get-WmiObject -Class Win32_Product | Where-Object -FilterScript { $_.name -like "*PCSC*" }
    
    # 2. ���oip
    $ipv4 = Get-IPv4Address 

    #����:
    #1�w�˪�������5.1.5.3 �� 51.5.5
    #2�ŦX���IP

    #���q��powershell v2�y�k, v2 ���䴩-in�y�k
    $installedVersions = "5.1.53", "5.1.55"
    $check_version = $false

    foreach ($version in $installedVersions) {
        if ($installedPCSC.version -eq $version) {
            $check_version = $true
            break
        }
    }

    $check_condition = $check_version -and ($ipv4 -like "172.20.*")


        #�g�J�O���ɸ��|
        $log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\pcsc_5157.log"
    
        $log_string = "Boot,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)"
        $log_string | Add-Content -PassThru $log_file
    
        if ($check_condition) {
            #�ŦX�ɯű���.
    
            #�����ª�
            $installedPCSC.uninstall()

            #�_��s��
            $new_pcsc_path = "\\172.20.5.187\mis\23-Ū�d������n��\CMS_CS5.1.5.7_20220925\CS5.1.5.7��_20220925"
    
            $new_pcsc_path = Get-Item $new_pcsc_path
            Copy-Item -Path $new_pcsc_path -Destination "C:\Vghtc\00_mis\$($new_pcsc_path.name)" -Recurse -Force -Verbose

            Stop-Process -Name csfsim -Force -ErrorAction SilentlyContinue

            #�w�˷s��
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\Vghtc\00_mis\$($new_pcsc_path.name)\gCIE_Setup\gCIE_Setup.msi /quiet /norestart" -Wait 

            #���s�Ұ��O�d�{��.
            Stop-Process -Name csfsim -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
            Start-Process -FilePath C:\nhi\BIN\csfsim.exe -ErrorAction SilentlyContinue

            #�O�����e
            $log_string = "Pass,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file


        }
        else {
            #���Ťɯű���
        
            if ($installedPCSC.Version -eq "5.1.57") { write-output "PCSC�����v�O5.1.57." }
            if ($ipv4.IPAddress -eq $null) { write-output "IP���D�W���|��IP." }
        
        }

    }    



    #�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q�פJ�ɤ��|����禡.
    if ($run_main -eq $null) {

        #�ˬd�O�_�޲z��
        $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

        if (!$check_admin -and !$runadmin) {
            #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
            Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
        }

        if ($check_admin) { 
            update-pcsc
        }
        else {
            Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
        }
        Write-Output "�ɯŧ���."
        Start-Sleep -Seconds 20
    }