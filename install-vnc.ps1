# �w��VNC
# setup.exe /? �i�H��ܦw�˫��O

# C:\ProgramData\Microsoft\Windows\Start Menu\Programs\UltraVNC\UltraVNC Server.lnk
# C:\ProgramData\Microsoft\Windows\Start Menu\Programs\UltraVNC\UltraVNC Viewer.lnk


param($runadmin)

$log_file = "\\172.20.1.14\update\0001-���a�t�����ҳ]�w\VHWC_logs\update-VNC.log"

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


# �w�q�Ыر��|�����
function New-Shortcut {
    param (
        [string]$TargetPath,
        [string]$ShortcutPath
    )
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.Save()
}

function install-VNC {
    
    $software_name = "UltraVnc"
    $software_path = "\\172.20.5.187\mis\08-VNC\1_4_36"
    # FIXME: $software_path = "\\172.20.1.122\share\software\00newpc\08-VNC\1_4_36"
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
            #Write-Log -logfile $log_file -message $msg
        
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

        #�w�˧�, �A���s���o�w�˸�T
        $all_installed_program = get-installedprogramlist
        $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }
    
    }

    #�]���o����MSI�w�˪�, �|�X�{�S���إߵ{�����|����, �ҥH��ʦh�إߤ@�Ӥ��ήୱ�����|

    $shortcuts = @{
        
        "viewer" = @{
            "name" = "UltraVNC Viewer.lnk"
            "folder" = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\UltraVNC"
            "exe" = "C:\Program Files\uvnc bvba\UltraVNC\vncviewer.exe"
        }

        "server" = @{
            "name" = "UltraVNC Server.lnk"
            "folder" = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\UltraVNC"
            "exe" = "C:\Program Files\uvnc bvba\UltraVNC\winvnc.exe"
        }
    
    }

    foreach ($shortcut in $shortcuts.keys) {

        $folderPath = $shortcuts.$shortcut.folder
        $shortcutPath = Join-Path $folderPath $shortcuts.$shortcut.name
        $exePath = $shortcuts.$shortcut.exe


        # �Ыظ�Ƨ��]�p�G���s�b�^
        if (!(Test-Path -Path $folderPath)) {
            try {
                New-Item -Path $folderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            catch {
                Write-Log -logfile $log_file -message "���~�G�L�k�Ыظ�Ƨ� $folderPath - $_"
                continue
            }
        }


        # �Ыر��|

        if (!(Test-Path -Path $shortcutPath)) {
            try {
                New-Shortcut -TargetPath $exePath -ShortcutPath $shortcutPath
                Write-Log -logfile $log_file -message "���\�Ыر��|�G$shortcutPath"
            }
            catch {
                Write-Log -logfile $log_file -message "���~�G�L�k�Ыر��| $shortcutPath - $_"
            }
         }
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


