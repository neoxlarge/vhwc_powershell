# �w��Java runtime 

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

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
                    
        #�����۰��ˬd��s
        New-Item -Path "HKCU:\SOFTWARE\JavaSoft\Java Update\Policy" -ItemType "Directory" -Force
        #New-ItemProperty -Path "HKCU:\SOFTWARE\JavaSoft\Java Update\Policy" -Name "EnableAutoUpdateCheck" -PropertyType "Binary" -Value "01 00 00 00 D0 8C 9D DF 01 15 D1 11 8C 7A 00 C0 4F C2 97 EB 01 00 00 00 C4 0D E3 D1 21 EC 3B 4B 99 6E CD 4B 90 BB BD 87 00 00 00 00 1C 00 00 00 50 00 61 00 73 00 73 00 77 00 6F 00 72 00 64 00 20 00 44 00 61 00 74 00 61 00 00 00 03 66 00 00 C0 00 00 00 10 00 00 00 A3 A8 42 B2 E9 B3 DE C6 27 52 6E 40 71 CB 34 53 00 00 00 00 04 80 00 00 A0 00 00 00 10 00 00 00 50 0F 30 FB EC 55 7D 4B 3E 66 F8 F1 F9 CA 76 A9 08 00 00 00 3A 0F EC 0C AC B5 35 4B 14 00 00 00 26 4B 3E A9 9F 3F E2 35 20 F8 53 F2 F4 47 14 F3 A1 F7 9F 27 "

               
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
    pause
}