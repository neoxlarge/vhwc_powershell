# �s������t�� 20240729�W�u
# �t�λݨD:
# HCAserversign ��ƤH���d��
# Hiicos �۵M�H��
#  desktop�񱶮|, ��chrome�w�˦�m
# ���X, �Ω�u�X��window

param($runadmin)

$log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VHWC_logs\install-2100_2nd.log"

function import-vhwcmis_module {
    $moudle_paths = @(
        if ($script:MyInvocation.MyCommand.Path) {"$(Split-Path $script:MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue)"},
        "d:\mis\vhwcmis",
        "c:\mis\vhwcmis",
        "\\172.20.5.185\powershell\vhwc_powershell",
        "\\172.20.1.14\share\00�s�q���w��\vhwc_powershell",
        "\\172.20.1.122\share\software\00newpc\vhwc_powershell",
        "\\172.19.1.229\cch-share\h040_�i�q��\vhwc_powershell"
    )

    $filename = "vhwcmis_module.psm1"

    foreach ($path in $moudle_paths) {
        
        if (Test-Path "$path\$filename") {
            write-output "$path\$filename"
            Import-Module "$path\$filename" -ErrorVariable $err_import_module
            if ($err_import_module -eq $null) {
                Write-Output "Imported module path successed: $path\$filename"
                break
            }
        }
    }

    $result = get-command -Name "get-installedprogramlist" -CommandType Function -ErrorAction SilentlyContinue
    if ($result -eq $null) {
        throw "�L�k���Jvhwcmis_module�Ҳ�, �{���L�k���`����. "
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
        if ($check_admin) {
            $proc =Start-Process powershell.exe -ArgumentList $argumentList -PassThru
        } else {
            $proc =Start-Process powershell.exe -ArgumentList $argumentList -Credential $credential -PassThru
        }
        
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
        if ($check_admin) {
            $proc = Start-Process powershell.exe -ArgumentList $argumentList -PassThru  
        } else {
            $proc = Start-Process powershell.exe -ArgumentList $argumentList -Credential $credential -PassThru
        }
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

            try {
                Start-Process -FilePath robocopy.exe -ArgumentList "$($env:temp)\ C:\Users\Public\Desktop\ ""�G�N����t��(Chrome).lnk""" -Credential $credential
            }
            catch {
                Copy-Item -Path $shortcutPath_temp -Destination "C:\Users\Public\Desktop\�G�N����t��(Chrome).lnk"
            }

            #Start-Process -FilePath robocopy.exe -ArgumentList "$($env:temp)\ C:\Users\Public\Desktop\ ""�G�N����t��(Chrome).lnk""" -Credential $credential
                        
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

        if ($check_admin) {
            Start-Process -FilePath robocopy.exe -ArgumentList "$($env:temp) C:\Users\Public\Desktop �G�N����t��(Chrome).lnk"
        } else {
            $credential = get-admin_cred
            Start-Process -FilePath robocopy.exe -ArgumentList "$($env:temp) C:\Users\Public\Desktop �G�N����t��(Chrome).lnk" -Credential $credential
        }
        
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