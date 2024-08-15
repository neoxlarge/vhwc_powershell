# �w�ˤu��{��7z
# 20242024 �v�令�|�۰ʲ����ª�����.

param($runadmin)


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
    
    if (! ($(Get-PSDrive -Name $nas_name -ErrorAction SilentlyContinue) -eq $null) ) {
        Remove-PSDrive -Name $nas_name
    }
    
}
import-vhwcmis_module

function install-7z {
    
    if (!$credential) {
        $credential = get-admin_cred
    }

    $software_name = "7-Zip*"
    $software_path = "\\172.20.5.187\mis\07-7-7z"
    $software_msi = "7z-x64.msi"
    $software_msi_x86 = "7z-x32.msi"

    ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $software_exec = $software_msi }
        "x86" { $software_exec = $software_msi_x86 }
        default { throw "$software_name �L�k���`�w��: ���䴩���t��:  $($env:PROCESSOR_ARCHITECTURE)" }
    }

    $msi_version = get-msiversion -MSIPATH $($software_path + "\" + $software_exec)
    Write-output "========================"
    Write-Output "Software: $software_name"
    Write-Output "Source path: $software_path\$software_exec"
    write-output "ource version: $msi_version"
    Write-output "========================"


    # �w���s�u��w�˨ӷ������|.
    if (!(Test-Path -Path $software_path)) {
        New-PSDrive -Name $software_name -Root "$software_path" -PSProvider FileSystem -Credential $credential
        }

    ## ��X�n��O�_�v�w��
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed) {
        #�v���w��
        # ��������s��
        $msi_version = get-msiversion -MSIPATH $($software_path + "\" + $software_exec)
        $check_version = compare-version -Version1 $msi_version -Version2 $software_is_installed.DisplayVersion

        if ($check_version) {
            #msi��������s,�����ª���, ���s�d��$software_is_installed
            Write-Output "Find installed version(old): $($software_is_installed.DisplayVersion)"
            Write-Output "Uninstall string: $($software_is_installed.uninstallString)"
            Write-Output "Uninstalling $software_name"
            Write-output "========================"
            uninstall-software -name $software_is_installed.DisplayName

            
            $all_installed_program = get-installedprogramlist
            $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
        }

    } 
    

    if ($null -eq $software_is_installed) {
    
        Write-Output "Insalling: $software_name"

        #�_���ɮר쥻���Ȧs"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        
        if ($software_exec -ne $null) {
            $msiExecArgs = "/i $($env:temp + "\" + $software_path.Name + "\" + $software_exec) /passive"
            
            if ($check_admin) {
                # ���޲z���v��
                $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiExecArgs -PassThru
            }
            else {
                # �L�޲z���v��
                $credential = get-admin_cred
                $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiExecArgs -Credential $credential -PassThru
            }
            
            $proc.WaitForExit()
        } 
      
        
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

    install-7z    

    pause
}