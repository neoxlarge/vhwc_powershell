# �w��VNC
# setup.exe /? �i�H��ܦw�˫��O

param($runadmin)

# import vhwcmis_module.psm1
# ���ovhwcmis_module.psm1��3�ؤ覡:
# 1.�{�������e���|, ���AD�W��Group police����,���|����e���|.
# 2.�`�Ϊ����|, d:\mis\vhwc_powershell, ���O�C�x������.
# 3.�s��NAS�W���o. �D���쪺�q���|�S��NAS���v��, ����ʳs�WNAS.

$log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VHWC_logs\update-VNC.log"

$pspaths = @()

$work_path = "$(Split-Path $script:MyInvocation.MyCommand.Path)\vhwcmis_module.psm1"
if (test-path -Path $work_path) { $pspaths += $work_path }

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


function install-VNC {
    
    $software_name = "UltraVnc"
    $software_path = "\\172.20.1.122\share\software\00newpc\08-VNC\1_4_36"
    $software_msi = "UltraVNC_X64.msi"
    $software_msi_x86 = "UltraVNC_X86.msi"
    $log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VHWC_logs\update-VNC.log"

    ## �P�_OS�O32(x86)�άO64(AMD64), ��L��(ARM64)���w��  
    switch ($env:PROCESSOR_ARCHITECTURE) {
        "AMD64" { $software_exec = $software_msi }
        "x86" { $software_exec = $software_msi_x86 }
        default { Write-Warning "Unsupport CPU or OS:"  $env:PROCESSOR_ARCHITECTURE; $software_exec = $null }
    }

    ## ��X�n��O�_�v�w��

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }


    if ($software_is_installed) {

        $msi_version = get-msiversion -MSIPATH "$software_path\$software_exec"
        $result = Compare-Version -Version1 $msi_version -Version2 $software_is_installed.DisplayVersion

        if ($result) {
            $ipv4 = Get-IPv4Address 
            Write-Log -logfile $log_file -message "Find old VNC version:$($software_is_installed.DisplayVersion)"
  
            Write-Output "Find old VNC $software_name, version: $($software_is_installed.DisplayVersion)"
            Write-Output "Removing old version."
            Start-Process -FilePath $software_is_installed.UninstallString -ArgumentList "/SILENT" -Wait
            $software_is_installed = $null

        } else {
            $msg = "Installed VNC: $($software_is_installed.DisplayVersion)"
            Write-Log -logfile $log_file -message $msg
        
        }
    }

    if ($null -eq $software_is_installed) {
    
        Write-Output "Start to insall: $software_name"

        #�_���ɮר쥻���Ȧs"
        $software_path = get-item -Path $software_path
        Copy-Item -Path $software_path -Destination $env:temp -Recurse -Force -Verbose

        <#
        SERVERVIEWER
        1 server
        2 viewer
        3 server + viewer
        SERVICE
        1 install
        2 not install

        PASSWORD = mypassword
        Sample
        UltraVNC_1436_X86.msi  SERVERVIEWER=3 SERVICE=1 PASSWORD="sysc0012"
        #>


        if ($null -ne $software_exec) {
            #Start-Process -FilePath $($env:temp + "\" + $software_path.Name + "\" + $software_exec) -ArgumentList "/silent /loadinf=$($env:temp + "\" + $software_path.Name + "\installvnc.inf")" -Wait
            $install_filepath = "$($env:temp)\$($software_path.Name)\$software_exec"

            $install_arrg = "/i $install_filepath /passive /norestart PASSWORD=""sysc0012"" SERVERVIEWER=3 SERVICE=1"
            
            Start-Process -FilePath "msiexec.exe" -ArgumentList $install_arrg -Wait

            
            Write-Log -logfile $log_file -message "Start to install VNC: $install_filepath"

            Start-Sleep -Seconds 5
        }
        else {
            Write-Warning "$software_name �L�k���`�w��."
        }
      
        #�_��]�w��vltravnc.ini ��C:\Program Files\uvnc bvba\UltraVNC
        Copy-Item -Path ($env:temp + "\" + $software_path.Name + "\ultravnc.ini") -Destination ($env:ProgramFiles + "\uvnc bvba\UltraVNC") -Force

        #�h�إߤ@�Ӥ��ήୱ�����|
        $shortcutPath = "c:\Users\Public\Desktop\VNC Viewer.lnk"
        $targetPath = "C:\Program Files\uvnc bvba\UltraVNC\vncviewer.exe"
        New-Item -ItemType SymbolicLink -Path $shortcutPath -Target $targetPath -ErrorAction SilentlyContinue

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
        install-VNC    
    }
    else {
        Write-Warning "�L�k���o�޲z���v���Ӧw�˳n��, �ХH�޲z���b������."
    }
    #pause
    Start-Sleep -Seconds 10
}


