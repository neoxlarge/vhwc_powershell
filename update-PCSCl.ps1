# �ɯŶ��ݦw�����󰷫O�dŪ�d������(PCSC 5.1.5.7)

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
    "wreh" = @{             #�_�ج�
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

function Check-EnvPathContains($path) {
    # �ˬd�t�������ܼ�$env:path�O�_�]�t�S�w���|
    $envPathList = $env:path -split ";"
    foreach ($p in $envPathList) {
        if ($p -like "*$path*") {
            #�t�������ܼƤ��]�t$path,���ݰ���C:\VGHTC\00_mis\���aiccard�����ܼƳ]�w.bat 
            return $true
        }
    }
    #�t�������ܼƤ����]�t$path,�ݰ���C:\VGHTC\00_mis\���aiccard�����ܼƳ]�w.bat 
    return $false
}


function Create-Shortcut {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath,

        [Parameter(Mandatory = $true)]
        [string]$ShortcutPath
    )

    # �Ы� WScript.Shell ����
    $shell = New-Object -ComObject WScript.Shell

    # �ˬd�ֱ��覡�O�_�w�s�b�A�p�G�s�b��R��
    if ($shell.CreateShortcut($ShortcutPath).FullName) {
        Remove-Item $ShortcutPath -Force -ErrorAction SilentlyContinue
    }

    # �Ыاֱ�
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Save()
    
}

function get-installedprogramlist {
    # ���o�Ҧ��w�˪��n��,���U�w�˳n��|�Ψ�.

    ### Win32_product���M��ä�����A Winnexus �ä��b�̭�.
    ### $all_installed_program = Get-WmiObject -Class Win32_Product

    ### �Ҧ����n��|�b���U�o�T�ӵn���ɸ��|��

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
        
        [string]$Token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI",                 # Line Notify �s���v��

        [Parameter(Mandatory = $true)]
        [string]$Message,               # �n�o�e���T�����e

        [string]$StickerPackageId,      # �n�@�ֶǰe���K�ϮM�� ID

        [string]$StickerId              # �n�@�ֶǰe���K�� ID
    )

    # Line Notify API �� URI
    $uri = "https://notify-api.line.me/api/notify"

    # �]�w HTTP Header�A�]�t Line Notify �s���v��
    $headers = @{ "Authorization" = "Bearer $Token" }

    # �]�w�n�ǰe���T�����e
    $payload = @{
        "message" = $Message
    }

    # �p�G�n�ǰe�K�ϡA�[�J�K�ϮM�� ID �M�K�� ID
    if ($StickerPackageId -and $StickerId) {
        $payload["stickerPackageId"] = $StickerPackageId
        $payload["stickerId"] = $StickerId
    }

    try {
        # �ϥ� Invoke-RestMethod �ǰe HTTP POST �ШD
        Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payload

        # �T�����\�ǰe
        Write-Output "�T���w���\�ǰe�C"
    }
    catch {
        # �o�Ϳ��~�A��X���~�T��
        Write-Error $_.Exception.Message
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
    $installedVersions = "5.1.51","5.1.53", "5.1.55"
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
    
    $log_string = "Boot,PCSC:$($installedPCSC.version),$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)"
    $log_string | Add-Content -PassThru $log_file
    Send-LineNotifyMessage -Message $log_string
    Start-Sleep -Seconds 2
    
    #�ŦX�ɯű���,�i��ɯ�
    if ($check_condition) {
                
        #�_��s��
        $new_pcsc_path = "\\172.20.5.187\mis\23-Ū�d������n��\CMS_CS5.1.5.7_20220925\CS5.1.5.7��_20220925"
    
        $new_pcsc_path = Get-Item $new_pcsc_path
        Copy-Item -Path $new_pcsc_path -Destination "C:\Vghtc\00_mis\$($new_pcsc_path.name)" -Recurse -Force -Verbose

        #�������O�d�n��
        Stop-Process -Name csfsim -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3

        #�����ª�
        $installedPCSC.uninstall()
        Start-Sleep -Seconds 3

        #�w�˷s��
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\Vghtc\00_mis\$($new_pcsc_path.name)\gCIE_Setup\gCIE_Setup.msi /quiet /norestart" -Wait 

        #�O�����e
        $log_string = "updated,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
        $log_string | Add-Content -PassThru $log_file
        Send-LineNotifyMessage -Message $log_string
        Start-Sleep -Seconds 2
    }
    elseif ( ($installedPCSC -eq $null) -and `
        (test-path C:\NHI\sam\COMX1\0640140012001000005984.SAM) ) {
        #�o�O�@�Ӵ��g�o�ͦ���]������bug, ���ݼҲճQ������]�����F. 
        #�ҥH�w�˦C�����s�b, ��SAM�o�s�b.
        #�o�{�o���δN�A�ˤ@��

        #�_��s��
        $new_pcsc_path = "\\172.20.5.187\mis\23-Ū�d������n��\CMS_CS5.1.5.7_20220925\CS5.1.5.7��_20220925"
    
        $new_pcsc_path = Get-Item $new_pcsc_path
        Copy-Item -Path $new_pcsc_path -Destination "C:\Vghtc\00_mis\$($new_pcsc_path.name)" -Recurse -Force -Verbose
       
        #�w�˷s��
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/i C:\Vghtc\00_mis\$($new_pcsc_path.name)\gCIE_Setup\gCIE_Setup.msi /quiet /norestart" -Wait 

        #�O�����e
        $log_string = "abnormal uninstall,PCSC:$($installedPCSC.version),$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
        $log_string | Add-Content -PassThru $log_file
        Send-LineNotifyMessage -Message $log_string
        Start-Sleep -Seconds 2
    }
    else {
        #���Ťɯű���
        
        if ($installedPCSC.Version -eq "5.1.57") { write-output "PCSC�����v�O5.1.57." }
        if ($ipv4 -eq $null) { write-output "IP���D�W���|��IP." }
    }

    #�ˬdSAM��
    #�ɯ����Ӥ����ˬd�o��.

    #�_��Link
    $diff1 = "C:\Users\Public\Desktop\���ݦw���ҲեD���x.lnk"
    $diff2 = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\���ݦw���ҲեD���x.lnk"
    $compare_result = Compare-Object -ReferenceObject (Get-Content $diff1) -DifferenceObject (Get-Content $diff2)

    if (($compare_result -ne $null) -and (Test-Path -Path $diff2)) {
        #�p�G�쥻�N���bstartup��Ƨ���, �]���δ_��L�h.
        Copy-Item -path $diff1 -Destination $diff2 -Force
        $log_string = "link copied,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
        $log_string | Add-Content -PassThru $log_file
        Send-LineNotifyMessage -Message $log_string
        Start-Sleep -Seconds 2
    }

    #�_��CSHIS.dll����w��Ƨ�

    $source_dll = Get-ItemProperty -Path "C:\NHI\LIB\CSHIS.dll" -ErrorAction SilentlyContinue

    if ($source_dll -ne $null) {
        
        $setup_file_ = @(
            "C:\VGHTC\ICCard\CSHIS.dll",
            "C:\ICCARD_HIS\CSHIS.dll",
            "C:\vhgp\ICCard\CSHIS.dll"
        )

        #�����U�C�{��,�H���e��DLL.
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
    #�N�O�_��"C:\NHI\LIB\"�̩Ҧ�dll��3�Ӹ�Ƨ�.
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

    #�ˬd�t�������ܼ�
    $setting_file = "C:\VGHTC\00_mis\���aiccard�����ܼƳ]�w.bat"
    Write-Output "�������ҳ]�w: $setting_file"
    $path = "C:\VGHTC\ICCard"
    $result = Check-EnvPathContains "C:\VGHTC\ICCard"

    if ($result -eq $false) {
    
        Write-Warning "�t�������ܼƤ����]�t $path"
        
        if ($check_admin) {
            if (Test-Path $setting_file) {
                Write-Output "����]�w��: $setting_file"
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c $setting_file"  -Wait
            }
        }
            
    }
    elseif ($result -eq $true) {

        Write-Output "�t�������ܼƤ��]�t$path"
        Write-Output "���ݰ��� $setting_file "

    }

}    


function update-virtualhc {
    #�ɯŵ������O�dSDK 2.5.4

    #�u�����˹L���~�ݭn�ɯ�
    #�s�����Φw��, �u�n�_���Ƨ��A����{���Y�i

    #���o�ª��n��
    $installed_vhc = get-installedprogramlist | Where-Object -FilterScript { $_.Displayname -like "�������O�d*" }

    $log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VirtualHC_254.log"

    if ($installed_vhc -ne $null) {
            
        #�_��2.5.4���쥻��
        $vhc_path = "\\172.20.5.187\mis\25-�������O�d\�E��\VHIC_virtual-nhicard+SDK+Setup-2.5.4"
        $vhc_path = Get-Item $vhc_path
        Copy-Item -Path $vhc_path -Destination "c:\NHI\$($vhc_path.name)" -Recurse -Force -Verbose

        #�����n��
        $unistalll_strign = $installed_vhc.QuietUninstallString.Split("""")
        Start-Process -FilePath $unistalll_strign[1] -ArgumentList $unistalll_strign[2] -Wait -ErrorAction SilentlyContinue -NoNewWindow
        #$installed_vhc.uninstall()
        Start-Sleep -Seconds 3

        #�_��|��ୱ�αҰ�
        Create-Shortcut -TargetPath "C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\�������O�d����n��-������.2.5.4.exe" -ShortcutPath "C:\users\public\desktop\�������O�d����n��.lnk"
        Create-Shortcut -TargetPath "C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\�������O�d����n��-������.2.5.4.exe" -ShortcutPath "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\�������O�d����n��.lnk"
    

        #open firewall
        #netsh advfirewall firewall add rule name='Allow �������O�d����n��' dir=in action=allow program='C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\�������O�d����n��-������.2.5.4.exe'
        Start-Process netsh.exe -ArgumentList "advfirewall firewall add rule name='Allow �������O�d����n��' dir=in action=allow program='C:\NHI\VHIC_virtual-nhicard+SDK+Setup-2.5.4\�������O�d����n��-������.2.5.4.exe'"

        $log_string = "update virtualHC$count,$env:COMPUTERNAME,$ipv4,$(Get-OSVersion),$env:PROCESSOR_ARCHITECTURE,$(Get-Date)" 
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
        update-virtualhc
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }

    #���s�Ұ��O�d�{��.
    Stop-Process -Name csfsim -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Start-Process -FilePath C:\nhi\BIN\csfsim.exe -ErrorAction SilentlyContinue

    Write-Output "�ɯŧ���."
    Start-Sleep -Seconds 20
}