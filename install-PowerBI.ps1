# �w��Winnexus

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


function install-PowerBI{
    # �w��Winnexus
    $software_name = "Microsoft PowerBI Desktop (x64)"
    $software_path = "\\172.20.5.187\mis\26-PowerBI"
    $software_exec = "PBIDesktopSetup_x64.exe"

    $Username = "vhcy\73058"
    $Password = "Q1220416+"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist

   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    if ($software_is_installed -eq $null) {
        Write-OutPut "Start to install: $software_name"

        $software_path_name = $software_path.Split("\")[-1]
        
        #�_���ɮר�temp
        #copy-item �L�k���{��, ���n�qpsdrive��, �ҥH�n��driver.
        $net_driver = "vhwcdrive" #�u�O����driver�W�r�Ӥv.
        New-PSDrive -Name $net_driver -Root $software_path -PSProvider FileSystem -Credential $credential
        Copy-Item -Path "$($net_driver):\*" -Destination "$($env:TEMP)\$software_path_name" -Force -Verbose
        Remove-PSDrive -Name $net_driver

        #installing...

        Start-Process -FilePath "$($env:TEMP)\$software_path_name\$software_exec" -ArgumentList "-passive -norestart ACCEPT_EULA=1" -Wait
        Start-Sleep -Seconds 5 
   
     
        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name *" }
    } 

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)



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
        install-PowerBI
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}