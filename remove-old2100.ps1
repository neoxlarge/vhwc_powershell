#remove old 2100

param($runadmin)

# import vhwcmis_module.psm1
# 取得vhwcmis_module.psm1的3種方式:
# 1.程式執行當前路徑, 放到AD上用Group police執行,不會有當前路徑.
# 2.常用的路徑, d:\mis\vhwc_powershell, 不是每台都有放.
# 3.連到NAS上取得. 非網域的電腦會沒有NAS的權限, 須手動連上NAS.

$log_file = "\\172.20.1.14\update\0001-中榮系統環境設定\VHWC_logs\remove-old2100.log"

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