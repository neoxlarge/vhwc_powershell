# �w��Java runtime 

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function set-Java_env {
    Write-Output "����JAVA�۰��ˬd��s"
    
    #�g�J���ȬObinary���A
    $binary_value = @(0x01, 0x00, 0x00, 0x00, 0xD0, 0x8C, 0x9D, 0xDF, 0x01, 0x15, 0xD1, 0x11,
        0x8C, 0x7A, 0x00, 0xC0, 0x4F, 0xC2, 0x97, 0xEB, 0x01, 0x00, 0x00, 0x00,
        0xC4, 0x0D, 0xE3, 0xD1, 0x21, 0xEC, 0x3B, 0x4B, 0x99, 0x6E, 0xCD, 0x4B,
        0x90, 0xBB, 0xBD, 0x87, 0x00, 0x00, 0x00, 0x00, 0x1C, 0x00, 0x00, 0x00,
        0x50, 0x00, 0x61, 0x00, 0x73, 0x00, 0x73, 0x00, 0x77, 0x00, 0x6F, 0x00,
        0x72, 0x00, 0x64, 0x00, 0x20, 0x00, 0x44, 0x00, 0x61, 0x00, 0x74, 0x00,
        0x61, 0x00, 0x00, 0x00, 0x03, 0x66, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00,
        0x10, 0x00, 0x00, 0x00, 0xA3, 0xA8, 0x42, 0xB2, 0xE9, 0xB3, 0xDE, 0xC6,
        0x27, 0x52, 0x6E, 0x40, 0x71, 0xCB, 0x34, 0x53, 0x00, 0x00, 0x00, 0x00,
        0x04, 0x80, 0x00, 0x00, 0xA0, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00,
        0x50, 0x0F, 0x30, 0xFB, 0xEC, 0x55, 0x7D, 0x4B, 0x3E, 0x66, 0xF8, 0xF1,
        0xF9, 0xCA, 0x76, 0xA9, 0x08, 0x00, 0x00, 0x00, 0x3A, 0x0F, 0xEC, 0x0C,
        0xAC, 0xB5, 0x35, 0x4B, 0x14, 0x00, 0x00, 0x00, 0x26, 0x4B, 0x3E, 0xA9,
        0x9F, 0x3F, 0xE2, 0x35, 0x20, 0xF8, 0x53, 0xF2, 0xF4, 0x47, 0x14, 0xF3,
        0xA1, 0xF7, 0x9F, 0x27)
    New-Item -Path "HKCU:\SOFTWARE\JavaSoft\Java Update\Policy" -ItemType "Directory" -Force
    New-ItemProperty -Path "HKCU:\SOFTWARE\JavaSoft\Java Update\Policy" -Name "EnableAutoUpdateCheck" -PropertyType "Binary" -Value $binary_value -Force
    
    Write-Output "�ק�JAVA����Ѽ�,�[�J-Xmx256m"
    <# 
    �n�۰ʭק� Java ����x�]�w�A�i�H�z�L PowerShell �ק� Java ����x�t�m��� deployment.properties�C
    �Ӥ��]�t�F Java ����x���U�ذt�m�ﶵ�A�z�i�H�z�L�ק�Ӥ��Ӧ۰ʤƧ�� Java ����x���]�w�C
    #>
    $filePath = "$env:USERPROFILE\AppData\LocalLow\Sun\Java\Deployment\deployment.properties"

    # �ˬd���O�_�s�b
    if (Test-Path $filePath) {
        # Ū����󤺮e
        $content = Get-Content $filePath -Raw
    
        # �N deployment.javaws.jre.0.args �ﶵ�]�m�� -Xmx256m
        $content = $content -replace "deployment.javaws.jre.0.args=.*", "deployment.javaws.jre.0.args=-Xmx256m"
    
        # �N�ק�᪺���e�g�^���
        Set-Content $filePath $content -Force
    }
    else {
        Write-Error "�䤣�� deployment properties ���C"
    }
}

function install-Java {

    $software_name = "Java*"
    $software_path = "\\172.20.5.187\mis\11-���a�t��\01-���a�t�����ҳ]�w\���a�t�����ҳ]�w\java"
    $software_msi = "jre-6u13-windows-i586-p.exe"

    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר쥻���Ȧs"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force

       
        Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_msi) -ArgumentList "/passive" -Wait
        Start-Sleep -Seconds 5 
                    
        
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
        install-Java    
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    
    set-Java_env
    pause
}