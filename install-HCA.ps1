## �w�� HCAServiSign

param($runadmin)

# �n�Dpowershell v5.1�H�W�~����, win7�w�]powershell v2.0.
if (!$PSVersionTable.PSCompatibleVersions -match "^5\.1") {
    Write-Output "powershell requires version 5.1, exit"
    Start-Sleep -Seconds 3
    exit
}

function import-vhwcmis_module {
    # import vhwcmis_module.psm1
    # ���ovhwcmis_module.psm1��3�ؤ覡:
    # 1.�{�������e���|, ���AD�W��Group police����,���|����e���|.
    # 2.�`�Ϊ����|, d:\mis\vhwc_powershell, ���O�C�x������.
    # 3.�s��NAS�W���o. �D���쪺�q���|�S��NAS���v��, ����ʳs�WNAS.

    $pspaths = @()

    if ($script:MyInvocation.MyCommand.Path -ne $null) {
        $work_path = "$(Split-Path $script:MyInvocation.MyCommand.Path)\vhwcmis_module.psm1"
        if (test-path -Path $work_path) { $pspaths += $work_path }
    }
    $nas_name = "nas122"
    $nas_path = "\\172.20.1.122\share\software\00newpc\vhwc_powershell"
    if (!(test-path $nas_path)) {
        $nas_Username = "software_download"
        $nas_Password = "Us2791072"
        $nas_securePassword = ConvertTo-SecureString $nas_Password -AsPlainText -Force
        $nas_credential = New-Object System.Management.Automation.PSCredential($nas_Username, $nas_securePassword)
        
        New-PSDrive -Name $nas_name -Root "$nas_path" -PSProvider FileSystem -Credential $nas_credential | Out-Null
    }
    $pspaths += "$nas_path\vhwcmis_module.psm1"

    $local_path = "d:\mis\vhwc_powershell\vhwcmis_module.psm1"
    if (Test-Path $local_path) { $pspaths += $local_path }

    foreach ($path in $pspaths) {
        Import-Module $path -ErrorAction SilentlyContinue
        if ((get-command -Name "get-installedprogramlist" -CommandType Function -ErrorAction SilentlyContinue)) {
            break
        }
    }
}
import-vhwcmis_module


function install-HCA {
    
    $log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VHWC_logs\update-HCA.log"

    ## �w�� HCAServiSign
    ### ��X�n��O�_�v�w��

    $software_name = "HCAServiSignAdapterSetup"
    $software_path = "\\172.20.1.122\share\software\00newpc\05-HCAServiSign��ƥd����"
    $software_exec = "HCAServiSignAdapterSetup.exe"

    $all_installed_program = get-installedprogramlist
  
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    if ($software_is_installed) {
        $exe_version = (Get-ItemProperty -Path "$software_path\$software_exec").VersionInfo.FileversionRaw.toString()
        $result = Compare-Version -Version1 $exe_version -Version2 $software_is_installed.DisplayVersion

        if ($result) {
            $ipv4 = Get-IPv4Address 
            Write-Log -logfile $log_file -message "Find old HCA version:$($software_is_installed.DisplayVersion)"
  
            Write-Output "Find old HCA $software_name, version: $($software_is_installed.DisplayVersion)"
            Write-Output "Removing old version."
            Start-Process -FilePath $software_is_installed.UninstallString -ArgumentList "/S" -Wait
            $software_is_installed = $null

        }
    }

    if ($null -eq $software_is_installed) {
        Write-Output "Start to install $software_name"

        #�ӷ����| ,�n�_����|,and �w�˰���{���W��
        $software_path = get-item -Path $software_path
        
        #�_���ɮר�Ȧs��Ƨ�
        Copy-Item -Path $software_path -Destination $env:TEMP -Recurse -Force -Verbose

        #installing...
        $process_id = Start-Process -FilePath ($env:TEMP + "\" + $software_path.Name + "\" + $software_exec) -PassThru
    
        #�̦w�ˤ��, HCAServiSignMonitor�|�̫�Q�}��, �ҥH�ˬd��ӵ{�ǰ����, ��ܦw�˧���.
        $process_exist = $null
        while ($process_exist -eq $null) {
            $process_exist = Get-Process -Name HCAServiSignMonitor -ErrorAction SilentlyContinue
            if ($process_exist -ne $null) { Stop-Process -Name $process_id.Name }
            write-output ($process_id.Name + "is installing...wait 5 seconds.")
            Start-Sleep -Seconds 5
        }


        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "$software_name*" }

    } 

    Write-output ("software has installed:" + $software_is_installed.DisplayName )
    Write-Output ("Version:" + $software_is_installed.DisplayVersion)

    if (! ($(Get-PSDrive -Name $nas_name -ErrorAction SilentlyContinue) -eq $null) ) {
        Remove-PSDrive -Name $nas_name
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

    if ($check_admin) { 
        install-HCA
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    #pause
    Start-Sleep -Seconds 5
}
