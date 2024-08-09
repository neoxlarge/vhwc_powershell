#remove old 2100

param($runadmin)

# import vhwcmis_module.psm1
# ���ovhwcmis_module.psm1��3�ؤ覡:
# 1.�{�������e���|, ���AD�W��Group police����,���|����e���|.
# 2.�`�Ϊ����|, d:\mis\vhwc_powershell, ���O�C�x������.
# 3.�s��NAS�W���o. �D���쪺�q���|�S��NAS���v��, ����ʳs�WNAS.

$log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VHWC_logs\remove-old2100.log"

$pspaths = @()
if ( $script:MyInvocation.MyCommand.Path -ne $null) {
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


$software_names = "�q�l����t��",
"UniView",
"IPD21",
"HiCOS PKI Smart Card Client v2.1.9.1u(with up2date)"

foreach ($software in $software_names) {
    
    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software }

    if ($software_is_installed -ne $null) {

        if ($software_is_installed.UninstallString -match "msiexec.exe") {
            uninstall-software -name $software_is_installed.DisplayName
            write-log -LogFile $log_file -Message "Removed  $($software_is_installed.DisplayName)"
           
        }
    }

}
