#�]�wadboe pdf���۰ʧ�s

param($runadmin)

<# 20230906
refer to https://www.adobe.com/devnet-docs/acrobatetk/tools/PrefRef/Windows/Updater-Win.html?zoom_highlight=updates#idkeyname_1_20396

�T���s���s
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown]
"bUpdater"=dword:00000000

0: Disables and locks the Updater.
1: No effect.


�T�Φ۰ʧ�s
�q��\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Adobe\Adobe ARM\Legacy\Reader\{AC76BA86-7AD7-1028-7B44-AC0F074E4100}
"Mode"=0

0: Do not download or install updates automatically.
1: Do not download or install updates automatically. Same as 0.
2: Automatically download updates but let the user choose when to install them.
3: Automatically download and install updates.
4: Notify the user downloads are available but do not download them.
#>


<# ���U��Ƥv���A��
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown]
"bUpdateToSingleApp"=dword:00000000

#�s��adobe reader��W��adobe acrobat, registry���H���ܧ�Adobe Arcobate.
#>

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-pdf {
    #�w��adobe pdf

    $software_name = "Adobe Acrobat*"
    $zip_path = "\\172.20.5.187\mis\16-PDF\adobe_pdf_���u�w����.zip"
    $software_path = $($zip_path.Split("\")[-1])
    $software_msi = "AcroRead.msi"
    $software_msi_update = "AcroRdrDCUpd2300120064.msp"

    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -eq $null) {

        if (Test-Path $zip_path) {
            Write-Output "Start to install $software_name"
            Write-Output "�_��B�����Y�ɮ�: $zip_path"
            expand-archive -Path $zip_path -DestinationPath $env:TEMP\$software_path -Force
        
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $env:TEMP\$software_path\$software_msi /passive /norestart" -wait
            Start-Sleep -Seconds 3
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/update $env:TEMP\$software_path\$software_msi_update /passive /norestart" -wait
            Start-Sleep -Seconds 3

        }
        else {
            Write-Warning "�w���ɸ��|���s�b,���ˬd,: $zip_path"
        }

        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    }

}

function check-pdf {

    $reg_path_disableUpdateButton= @("HKLM:\SOFTWARE\Policies\Adobe\Adobe Acrobat\DC\FeatureLockDown",
        "HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown")

    $reg_path_disableAutoupdate = "HKLM:\SOFTWARE\WOW6432Node\Adobe\Adobe ARM\Legacy\Reader\{AC76BA86-7AD7-1028-7B44-AC0F074E4100}"

    foreach ($r in $reg_path_disableUpdateButton) {
        if (Test-Path -Path $r) {
            $is_update = Get-ItemProperty -Path $r -Name "bUpdater" -ErrorAction SilentlyContinue

            if ($is_update.bUpdateToSingleApp -eq 0) {
                Write-Output "Adobe PDF Reader �v�]�w���۰ʧ�s."
            }
            else {
                if ($check_admin) {
                    Write-Output "���b�]�wAdobe PDF Reader �����۰ʷs��."
                    Set-ItemProperty -Path $r -Name "bUpdater" -Value 00000000 -type DWord -Force
                    Set-ItemProperty -Path $reg_path_disableAutoupdate -Name "Mode" -Value 0 -Force -type Dword
                }
                else {
                    Write-Warning "�S���t�κ޲z���v��,�BAdobe PDF Reader���]�w���۰ʷs,�ХH�t�κ޲z���������s����."
                }
            }    
        }
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
    else {
        install-pdf
    }
    
    check-pdf
    pause
}