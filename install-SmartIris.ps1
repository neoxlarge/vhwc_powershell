# �w��SmartIris

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install-SmartIris {
    
    $software_name = "SmartIris"
    $software_path = "\\172.20.5.187\mis\02-SmartIris\SmartIris_V1.3.6.4_Beta7_UQ-1.1.0.19_R2_Install_20200701"
    $software_exe = "setup.exe"
    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר쥻���Ȧs"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        # �w��  
        $runid = Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exe) -ArgumentList "/s /f1$($env:temp + "\" + $software_path.Name + "\vhwc.iss")" -PassThru
        
        # �w�˹L�{��, �o2��{�|���X��, ����F, ���|�A�]�w�Y�i.
        while (!($runid.HasExited)) {
            get-process -Name MonitorCfg -ErrorAction SilentlyContinue | Stop-Process
            get-process -Name UQ_Setting -ErrorAction SilentlyContinue | Stop-Process
            Start-Sleep -Seconds 1
        }

        #�̷� \\172.19.1.14\Update\��T��\�@�ε{��\SmartIris\SmartIris��SOP.doc
        #�B�J�T�G
        #\\172.19.1.14\Update\��T��\�@�ε{��\SmartIris\UltraQuery_V1.1.1.0_Update_20200731
        # ����ƻs��C:\TEDPC\SmartIris\UltraQuery���л\

        copy-item -Path "\\172.20.5.187\mis\02-SmartIris\UltraQuery_V1.1.1.0_Update_20200731\*" -Destination "C:\TEDPC\SmartIris\UltraQuery" -Recurse -Force

        #�_��]�w�ɨ쥻��.
        copy-item -Path "\\172.20.5.187\mis\02-SmartIris\vhwc_UltraQuery_SysIni\*" -Destination "C:\TEDPC\SmartIris\UltraQuery\SysIni" -Force
 

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
        install-SmartIris    
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}