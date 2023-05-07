# �w�ˤ���t��2100
# �w�˱оǤ�� http://172.22.250.179/ii/install/sysdeploy.htm


param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

function install_msi {
    #$mode �Omsiexec���Ѽ�, �w�]i�O�w��, fa�O�j��s��
    #msi�O�w�˪��ɦW
    param($mode = "i", $msi)
    Write-Output $msi
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/$mode $msi /passive /norestart" -wait

    Start-Sleep -Seconds 3
}


function install-2100 {

    $software_name = "�q�l����t��"
    $software_path = "\\172.20.5.187\mis\08-2100����t��\01.2100����t�Φw�˥]_Standard"
    
    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    #�_���ɮר쥻���Ȧs"
    $software_path = get-item -Path $software_path
    Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force

    #�n�w�˪��ɮ�
    $package_msi = @(
        "eDocSetup_Win7.msi",
        #"HC_Setup.msi", HCA��ƤH�������X�ʵ{��(���ѮH�����|�w��), �o����]�|��
        #"HiCOS.msi", #HiCOP����|�A��, �o�̤���
        "IPD21XSetup.msi",
        "soapsdk.msi",
        #"SetupXP.msi",  #XP�Ϊ����θ�
        "UniView.msi"
    )


    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"
        #�w��MSI��
        foreach ($p in $package_msi) {
            $p_path = $env:temp + "\" + $software_path.Name + "\package\" + $p 
            install_msi -mode "i" -msi $p_path
        }

        #�w��tablePC_SDK.exe
        Start-Process -FilePath ($env:temp + "\" + $software_path.Name + "\package\tablePC_SDK.exe") -ArgumentList "/v/passive" -Wait

    }
    else {
        Write-Output "Reinstall $software_name"

        foreach ($p in $package_msi) {
            $p_path = $env:temp + "\" + $software_path.Name + "\package\" + $p 
            install_msi -mode "fa" -msi $p_path
        }

        #�w��tablePC_SDK.exe
        Start-Process -FilePath ($env:temp + "\" + $software_path.Name + "\package\tablePC_SDK.exe") -ArgumentList "/v/passive" -wait

    }

    #�w�˧�, �A���s���o�w�˸�T
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }


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
        install-2100
        Import-Module ((Split-Path $PSCommandPath) + "\Check-2100env.ps1")    
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    #pause
}