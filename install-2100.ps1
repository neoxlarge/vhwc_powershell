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

function set-2100_env {

    #1.
    Write-Output "�_����|�]�w��ClientSetting.ini"    
    $path = $env:TEMP + "\01.2100����t�Φw�˥]_Standard\ClientSetting\ClientSetting_Chiayi.ini"
    if (Test-Path -Path $path) {
        copy-item -Path $path -Destination $env:SystemDrive\2100\SSO\ClientSetting.ini -Force
    } else {
        Write-Warning "�䤣��ClientSetting_Chiayi.ini�ɮ�, ���ˬd!!"
    }

    #2.
    <#
    hicos3.1������A�|�L�kñ�֤��媺�ѨM��k�G
    �Ч�U�C�o���ɮ�HiCOSCSPv32��b�좵���줸�q��C:\Windows\SysWOW64�A�����줸�q��C:\Windows\System32�A�t�~��.1�d���޲z�u��̪��u�]�w�v�����ӳ����ġA�A����@�U�����ɡA�Y�i�ѨM
    Hicoscspv32.dll ����: 
    #>

    switch ($env:PROCESSOR_ARCHITECTURE) {
        "amd64" {$dll_path = "$env:windir\SysWoW64"}
        "x86" {$dll_path = "$env:windir\System32"}
        default {Write-Warning "Unknown processor architecture."}
    }

    $dll = Get-ItemPropertyValue -Path "$dll_path\HiCOSCSPv32.dll" -Name "VersionInfo"
    if ($dll.ProductVersion -ne "3.0.3.21207") {
        if (Test-Path -Path ($env:temp + "\01.2100����t�Φw�˥]_Standard\HiCOSCSPv32.dll")) {
            #�л\Hicoscspv32.cll��c:\windows\system32��.
            Write-Output "�л\Hicoscspv32.cll(3.0.3.21207)��$dll_path"
            copy-item -Path ($env:temp +"\01.2100����t�Φw�˥]_Standard\HiCOSCSPv32.dll") -Destination $dll_path -Force

        } else {write-warning "�䤣�쥿�T��HiCOSCSPv32.dll�ɮ�"}
    }

    #3.
    Write-Output "���� 01����������.exe ��IE �]�w"
    Start-Process -FilePath reg.exe -ArgumentList ("import " + $env:temp + "\01.2100����t�Φw�˥]_Standard\reg\IE9setting.reg") -Wait
    Start-Process -FilePath reg.exe -ArgumentList ("import " + $env:temp + "\01.2100����t�Φw�˥]_Standard\reg\IE9setting1.reg") -Wait
    Start-Process -FilePath ($env:temp + "\01.2100����t�Φw�˥]_Standard\01����������.exe") -Wait
  

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
        #Import-Module ((Split-Path $PSCommandPath) + "\Check-2100env.ps1")    
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }

    set-2100_env
    pause
}