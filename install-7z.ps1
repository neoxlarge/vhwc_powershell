# �w�ˤu��{��7z
# 20242024 �v�令�|�۰ʲ����ª�����.

param($runadmin)

Import-Module -name "$(Split-Path $PSCommandPath)\vhwcmis_module.psm1"

function install-7z {
    
    $software_name = "7-Zip*"
    $software_path = "\\172.20.5.187\mis\07-7-7z"
    $software_msi = "7z2201-x64.msi"
    $software_msi_x86 = "7z2201-x32.msi"

    ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $software_exec = $software_msi }
        "x86" { $software_exec = $software_msi_x86 }
        default { throw "$software_name �L�k���`�w��: ���䴩���t��:  $($env:PROCESSOR_ARCHITECTURE)" }
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
            Write-Output "����ª�����: $($software_is_installed.DisplayName) : $($software_is_installed.DisplayVersion)"
            Write-Output "Uninstall string: $($software_is_installed.uninstallString)"
            uninstall-software -name $software_is_installed.DisplayName
            
            $all_installed_program = get-installedprogramlist
            $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
        }

    } 
    

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

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