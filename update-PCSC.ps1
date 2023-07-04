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
    
        #�ŦX�ɯű���,�i��ɯ�
        if ($check_condition) {
                
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
            $log_string = "updated,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file
        }
        else {
            #���Ťɯű���
        
            if ($installedPCSC.Version -eq "5.1.57") { write-output "PCSC�����v�O5.1.57." }
            if ($ipv4.IPAddress -eq $null) { write-output "IP���D�W���|��IP." }
        }

        #�ˬdSAM��
        #�ɯ����Ӥ����ˬd�o��.

        #�_��Link
        $diff1 = "C:\Users\Public\Desktop\���ݦw���ҲեD���x.link"
        $diff2 = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\���ݦw���ҲեD���x.link"
        $compare_result = Compare-Object -ReferenceObject (Get-Content $diff1) -DifferenceObject (Get-Content $diff2)

        if ($compare_result -ne $null) {
            Copy-Item -path $diff1 -Destination $diff2 -Force
            $log_string = "link copied,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
            $log_string | Add-Content -PassThru $log_file
        }

        #�_��CSHIS.dll����w��Ƨ�

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
        #�N�O�_��"C:\NHI\LIB\"�̩Ҧ�dll��3�Ӹ�Ƨ�.
        $setup_file_ = Get-ChildItem -Path "C:\NHI\LIB\"
    
        $setup_file_target_path = @(
            "C:\ICCARD_HIS",
            "C:\Windows\System32",
            "C:\Windows\System"    
        )

        $count = 0

        foreach ($i in $setup_file_) {
            Write-Output ("dll name: " + $i.Name + "dll versoin: " + $i.VersionInfo.ProductVersion    )

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