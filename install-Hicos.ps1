# �w��Hicos
# ��Jsetup.exe /?, �i�H�ݨ�䴩silent�w�˪����R
# /install /repair /uninstall 
# /passive /quiet
# /norestart /log log.txt

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-HiCOS {


    $software_name = "HiCOS*"
    $software_path = "\\172.20.5.187\mis\04-�۵M�H����-�s��HiCOS\HiCOS_Client-3.1.0.22133-20220624"
    $software_exec = "HiCOS_Client.exe"
    #$software_msi_x86 = "\EZ100_Driver_32bit\setup.exe"

    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    ## ����t�η|�w��2.1.9.1, �����o��

    foreach ($software in $software_is_installed) {
        if ($software.DisplayVersion -eq "2.1.9.1") {
            $uninstallstring = $software.uninstallString.Split(" ")[1].replace("/I", "/x")
            Start-Process -FilePath "msiexec.exe" -ArgumentList "$uninstallstring /passive" -Wait
            Start-Sleep -Seconds 5
            $software_is_installed = $software_is_installed | Where-Object { $_.DisplayVersion -ne "2.1.9.1" } #�q�}�C������
        }
    }


    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר쥻���Ȧs"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force

        $pid_ = Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/install /passive" -PassThru #����������-wait�|�����U��.
        while (!($pid_.HasExited)) { Start-Sleep -Seconds 5}
                    
        
        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

        #change HiCOS setting
        
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" {
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card V3 (Conf.1 T=CL)" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card V3 (Conf.1 T=1)" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card 32K" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Cryptography\Calais\SmartCards\CHT ePKI Card 32K" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\WOW6432Node\Chunghwa TeleCom\HiCOS PKI Smart Card\TokenUtility\1.0.0" -Name "CertImportWithLegacyCSP" -Value 1 -Force -PropertyType "DWord"
                Set-ItemProperty -Path "HKLM:\Software\WOW6432Node\Chunghwa TeleCom\HiCOS PKI Smart Card\TokenUtility\1.0.0" -Name "LegacyCSPParing" -Value 1 -Force
            }
            "x86" {
                New-ItemProperty -Path "HKLM:\Software\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card V3 (Conf.1 T=CL)" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card V3 (Conf.1 T=1)" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\Microsoft\Cryptography\Calais\SmartCards\CHT GPKI Card 32K" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\Microsoft\Cryptography\Calais\SmartCards\CHT ePKI Card 32K" -Name "Crypto Provider" -Value "HiCOS PKI Smart Card Cryptographic Service Provider" -Force -PropertyType "string"
                New-ItemProperty -Path "HKLM:\Software\Chunghwa TeleCom\HiCOS PKI Smart Card\TokenUtility\1.0.0" -Name "CertImportWithLegacyCSP" -Value 1 -Force -PropertyType "DWord"
                Set-ItemProperty -Path "HKLM:\Software\Chunghwa TeleCom\HiCOS PKI Smart Card\TokenUtility\1.0.0" -Name "LegacyCSPParing" -Value 1 -Force
            }
        }
    
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
        install-HiCOS
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}