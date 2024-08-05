# �s������t�� 20240729�W�u
# �t�λݨD:
# HCAserversign ��ƤH���d��
# Hiicos �۵M�H��
#  desktop�񱶮|, ��chrome�w�˦�m
# ���X, �Ω�u�X��window

param($runadmin)

$log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VHWC_logs\install-2100(chrome).log"

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



# ��ơG��� Chrome ����ڦw�˸��|
function Get-ChromePath {
    $chromePaths = @(
        (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"),
        (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe")
    )

    foreach ($path in $chromePaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    Write-Output "�L�k��� Chrome�C�нT�O Chrome �w�w�ˡC"
    exit
}

function Update-RegistryKey($keyPath, $valueName, $desiredValue) {
    # �Ыؤ@�ӷs�� PSDrive �ӳX�� HKLM
    New-PSDrive -Name HKLM -PSProvider Registry -Root HKEY_LOCAL_MACHINE -ErrorAction SilentlyContinue | Out-Null

    $fullPath = "HKLM:\$keyPath"
    if (-not (Test-Path $fullPath)) {
        New-Item -Path $fullPath -Force | Out-Null
        Write-Output "�إ߷s�����U��: $fullPath"
    }

    $currentValue = Get-ItemProperty -Path $fullPath -Name $valueName -ErrorAction SilentlyContinue
    if ($currentValue -eq $null -or $currentValue.$valueName -ne $desiredValue) {
        New-ItemProperty -Path $fullPath -Name $valueName -Value $desiredValue -PropertyType String -Force | Out-Null
        Write-Output "��s���U��: $fullPath\$valueName"
    }
    else {
        Write-Output "���U���w�s�b�B���T: $fullPath\$valueName"
    }

    # ���� PSDrive
    Remove-PSDrive -Name HKLM -ErrorAction SilentlyContinue
}



function install-2100_chrome() {

    $credential = get-admin_cred

    # ��ơG��� Chrome ����ڦw�˸��|


    # ��� Chrome ���|
    $chromePath = Get-ChromePath

    # �]�w���|���|
    $shortcutPath = Join-Path ([System.Environment]::GetFolderPath("CommonDesktopDirectory")) "�G�N����t��(Chrome).lnk"

    # �ˬd���|�O�_�w�s�b
    if (Test-Path $shortcutPath) {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
    
        if ($shortcut.TargetPath -ne $chromePath) {
            Write-Output "���|�w�s�b�A�� Chrome ���|�����T�C���b��s..."
            $shortcut.TargetPath = $chromePath
            $shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
            $shortcut.Save()
            Write-Output "���|�w��s�C"
        }
        else {
            Write-Output "���|�w�s�b�B Chrome ���|���T�C�L�ݧ��C"
        }
    }
    else {
        # �إ߷s���|
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = $chromePath
        $Shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
        $Shortcut.Save()

        Write-Output "�s���| '�G�N����t��(Chrome).lnk' �w�إߧ����C"
    }

    # �ˬd�M��s���U��
    $chromeKeyPath = "SOFTWARE\Policies\Google\Chrome\PopupsAllowedForUrls"
    $edgeKeyPath = "SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"
    $desiredValue = "edap.doc.vghtc.gov.tw"



    Update-RegistryKey $chromeKeyPath "99999" $desiredValue
    Update-RegistryKey $edgeKeyPath "99999" $desiredValue

    Write-Output "���U���ˬd�M��s�����C"

}





#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q���J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    install-2100_chrome    

    #pause
    Start-Sleep -Seconds 10
}