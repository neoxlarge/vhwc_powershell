# �s������t�� 20240729�W�u
# �t�λݨD:
# HCAserversign ��ƤH���d��
# Hiicos �۵M�H��
#  desktop�񱶮|, ��chrome�w�˦�m
# ���X, �Ω�u�X��window

param($runadmin)

$log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VHWC_logs\install-2100_2nd.log"

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
            
    # �ˬd�óЫص��U���]�p�G���s�b�^
    if (-not (Test-Path $keypath)) {
        $createItemCode = {
            param($path)
            New-Item -Path $path -Force
        }

        $scriptString = $createItemCode.ToString()
        $argumentList = "-ExecutionPolicy Bypass -Command `"& {$scriptString} -path '$keypath'`""

        $proc =Start-Process powershell.exe -ArgumentList $argumentList -Credential $credential -PassThru
        $proc.WaitForExit()
        Write-Output "�إ߷s�����U��: $keypath"
    }

    # �ˬd���U��ȬO�_�s�b�B���T
        $currentValue = Get-ItemProperty -Path $keypath -Name $valuename -ErrorAction SilentlyContinue
        if ($currentValue -eq $null -or $currentValue.$valuename -ne $desiredValue) {
            $needsUpdate = $true
        } else {
            $needsUpdate = $false
        }
   
   
    if ($needsUpdate) {
        # ��s���U���
        $updatePropertyCode = {
            param($path, $name, $value)
            New-ItemProperty -Path $path -Name $name -Value $value -PropertyType String -Force
        }

        $scriptString = $updatePropertyCode.ToString()
        $argumentList = "-ExecutionPolicy Bypass -Command `"& {$scriptString} -path '$keypath' -name '$valueName' -value '$desiredValue'`""
        $proc = Start-Process powershell.exe -ArgumentList $argumentList -Credential $credential -PassThru
        $proc.WaitForExit()
       
        Write-Output "�G�N����t�Χ�s���U��: $keypath\$valueName"
        Write-Log -LogFile $log_file -Message "�G�N����t�Χ�s���U��: $keypath\$valueName"
    }
    else {
        Write-Output "�G�N����t�ε��U���w�s�b�B���T: $keypath\$valueName"
    }
}



function install-2100_2nd() {

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
            #Write-Output "�G�N����t�α��|�w�s�b�A�� Chrome ���|�����T�C���b��s..."
            #$shortcut.TargetPath = $chromePath
            #$shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
            #$shortcut.Save()
            
            $shortcutPath_temp = Join-Path $env:temp "�G�N����t��(Chrome).lnk"

            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($shortcutPath_temp)
            $Shortcut.TargetPath = $chromePath
            $Shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
            $Shortcut.Save()

            $credential = get-admin_cred
            Start-Process -FilePath robocopy.exe -ArgumentList "$($env:temp) C:\Users\Public\Desktop �G�N����t��(Chrome).lnk" -Credential $credential
            
            
            Write-Output "�G�N����t�α��|�w��s�C"
            Write-log -LogFile $log_file -Message "�즳�G�N����t�α��|���e���~,���|�w��s�C  "
        }
        else {
            Write-Output "�G�N����t�α��|�w�s�b�B Chrome ���|���T�C�L�ݧ��C"
        }
    }
    else {
        # �إ߷s���|
        $shortcutPath_temp = Join-Path $env:temp "�G�N����t��(Chrome).lnk"

        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath_temp)
        $Shortcut.TargetPath = $chromePath
        $Shortcut.Arguments = "https://edap.doc.vghtc.gov.tw/ms/SSO.html"
        $Shortcut.Save()

        $credential = get-admin_cred
        Start-Process -FilePath robocopy.exe -ArgumentList "$($env:temp) C:\Users\Public\Desktop �G�N����t��(Chrome).lnk" -Credential $credential
        Write-Output "�s���| '�G�N����t��(Chrome).lnk' �w�إߧ����C"
        Write-log -LogFile $log_file -Message "�s���| '�G�N����t��(Chrome).lnk' �w�إߧ����C"
    }

    # �ˬd�M��s���U��
    $chromeKeyPath = "HKLM:\SOFTWARE\Policies\Google\Chrome\PopupsAllowedForUrls"
    $edgeKeyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"
    $desiredValue = "edap.doc.vghtc.gov.tw"


    Update-RegistryKey $chromeKeyPath "99999" $desiredValue
    Update-RegistryKey $edgeKeyPath "99999" $desiredValue

    Write-Output "�G�N����t�ε��U���ˬd�M��s�����C"

}





#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q���J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    install-2100_2nd    

    #pause
    Start-Sleep -Seconds 10
}