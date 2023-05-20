# �w��Java runtime 

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function set-Java_env {
    Write-Output "����JAVA�۰��ˬd��s"
    
    #�Ѧҳs�� https://thegeekpage.com/turn-off-java-update-notification-in-windows-10/
    #32bit/64bit�t�έn�ק諸registry��m���@��.
    #�ק勵�T�|���s�������������������.
    #�S��admin�v�����ΰ��������s�ʧ@.

    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $reg_path = "HKLM:\SOFTWARE\WOW6432Node\JavaSoft\Java Update\Policy" }
        "x86" { $reg_path = "HKLM:\SOFTWARE\JavaSoft\Java Update\Policy" }
    }
    
    
    if ((Get-ItemProperty -Path $reg_path -Name "EnableJavaUpdate").enableJavaupdate -ne 0 ) 
    {
        if ($check_admin) {
            Set-ItemProperty -Path $reg_path -Name "EnableJavaUpdate" -Value 0 -Force
        } else {
            Write-Warning "Java�v�ҥΦ۰ʧ�s, ���S���޲z���v���ק�, �Хκ޲z�̭���."
        }
    }  else {
        Write-Output "Java��s�v����"
    }

    <# ������ӧ�s������,���A�ݭn�H�U�]�w.
    Write-Output "�ק�Java��s�q�����U�����e"
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" {
            Set-ItemProperty -Path "HKLM:\Software\WOW6432Node\JavaSoft\Java Update\Policy" -name "NotifyInstall" -Value 0 -Force
            Set-ItemProperty -Path "HKLM:\Software\WOW6432Node\JavaSoft\Java Update\Policy" -name "NotifyDownload" -Value 1 -Force
        }
    
        "x86" {
            Set-ItemProperty -Path "HKLM:\Software\JavaSoft\Java Update\Policy" -name "NotifyInstall" -Value 0 -Force
            Set-ItemProperty -Path "HKLM:\Software\JavaSoft\Java Update\Policy" -name "NotifyDownload" -Value 1 -Force
        }
    }
    #>


    Write-Output "�ק�JAVA����Ѽ�,�[�J-Xmx256m"
    <# 
    �n�۰ʭק� Java ����x�]�w�A�i�H�z�L PowerShell �ק� Java ����x�t�m��� deployment.properties�C
    �Ӥ��]�t�F Java ����x���U�ذt�m�ﶵ�A�z�i�H�z�L�ק�Ӥ��Ӧ۰ʤƧ�� Java ����x���]�w�C
    #>
    $filePath = "$env:USERPROFILE\AppData\LocalLow\Sun\Java\Deployment\deployment.properties"

    #�b��w�˧�JAVA��, ������JAVA����x��, deployment.properties�ɮ׷|���b, ����JAVA����x�|���͸���.
    if (!(test-path $filePath)) {
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" {Start-Process -FilePath "C:\Program Files (x86)\Java\jre6\bin\javacpl.exe"}
            "x86" {Start-Process -FilePath "C:\Program Files\Java\jre6\bin\javacpl.exe"}
        }
        Start-Sleep -Seconds 5
        Stop-Process -Name "javaw"
    }

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