# install LibreOffice
# 20242024 �v�令�|�۰ʲ����ª�����.


param($runadmin)

$mymodule_path = Split-Path $PSCommandPath + "\"
Import-Module $mymodule_path + "get-installedprogramlist.psm1"
Import-Module $mymodule_path + "get-msiversion.psm1"
Import-Module $mymodule_path + "compare-version.psm1"
Import-Module $mymodule_path + "get-admin_cred.psm1"

function install-libreoffice {
    
    $software_name = "LibreOffice*"
    $software_path = "\\172.20.1.122\share\software\00newpc\10-LibreOffice"
    $software_msi = "LibreOffice_Win_x64.msi"
    $software_msi_x86 = "LibreOffice_Win_x86.msi"


    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed) {
        #�v���w��
        # ��������s��
        $msi_version = get-msiversion -MSIPATH ($software_path + "\" + $software_exec)
        $check_version = compare-version -Version1 $msi_version -Version2 $software_is_installed.DisplayVersion

        if ($check_version) {
            #msi��������s,�����ª���, ��$software_is_installed�M��
            Write-Output "����ª�����: $($software_is_installed.DisplayName) : $($software_is_installed.DisplayVersion)"
            uninstall-software -name $software_is_installed.DisplayName
            $software_is_installed = $null
        }

    } 
    
    
    if ($software_is_installed -eq $null) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר쥻���Ȧs"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
        switch ($env:PROCESSOR_ARCHITECTURE) {
            "AMD64" { $software_exec = $software_msi }
            "x86" { $software_exec = $software_msi_x86 }
            default { Write-Warning "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
        }

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
        else {
            Write-Warning "$software_name �L�k���`�w��."
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
    
    install-libreoffice
    
    pause
}