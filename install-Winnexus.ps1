# �w��Winnexus

param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")


function install-WinNexus {
    # �w��Winnexus
    $software_name = "WinNexus"
    #$software_path = "\\172.20.5.187\mis\13-Winnexus\Winnexus_1.2.4.7"
    $software_path = "\\172.20.1.122\share\software\00newpc\13-Winnexus"
    $software_exec = "Install_Desktop.1.2.4.7.exe"

    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist

   
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    if ($software_is_installed -eq $null) {
        Write-OutPut "Start to install: $software_name"

        $software_path = get-item -Path $software_path
        
        #�_���ɮר�D:\mis
        Copy-Item -Path $software_path -Destination $env:TEMP -Recurse -Force -Verbose

        #installing...
        Start-Process -FilePath ($env:TEMP + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList ("/suppressmsgboxes /log:install_winnexus.log") -Wait
        Start-Sleep -Seconds 5 
   
     
        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name *" }
    } 

    Write-Output ("Software has installed: " + $software_is_installed.DisplayName)
    Write-Output ("Version: " + $software_is_installed.DisplayVersion)



}



#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q?�J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    if ($check_admin) { 
        install-WinNexus
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    pause
}