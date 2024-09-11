#�����n���(Win10)
param($runadmin)

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


function Remove-AppsinWin10 {
    # �n���������ε{���M�� (win10 only)
    $appsToRemove = @(
        "Microsoft.SkypeApp", 
        "Microsoft.OneDrive", 
        "Microsoft.MicrosoftOfficeHub",
        #"Microsoft.XboxIdentityProvider",   #�����ε{���O Windows ���@�����A�ӥB�L�k�w��ӧO�ϥΪ̸Ѱ��w�ˡC
        #"Microsoft.XboxGameCallableUI",     #�����ε{���O Windows ���@�����A�ӥB�L�k�w��ӧO�ϥΪ̸Ѱ��w�ˡC
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.XboxGamingOverlay",      
        "Microsoft.XboxGameOverlay",        
        "Microsoft.XboxApp",                
        "Microsoft.Xbox.TCUI",
        "5A894077.McAfeeSecurity",
        "B9ECED6F.ASUSPCAssistant",
        "4DF9E0F8.Netflix",
        "ZhuhaiKingsoftOfficeSoftw.WPSOffice")              

    # �M���C�ӭn���������ε{��
    foreach ($app in $appsToRemove) {
        # �ˬd�O�_�w�g�w�˸����ε{��
        if (Get-AppxPackage -Name $app) {
            # �������ε{��
            Remove-AppxPackage -Package $(Get-AppxPackage -Name $app)
            Write-output "�w���\���� $app ���ε{���C"
        }
        else {
            Write-output "$app ���ε{�����w�ˡC"
        }
    }
}


function update-apps {

    # �]�w $DebugPreference �� "Continue" ��ܰT��.
    $DebugPreference = "Continue"
    $apps = @{

        "StickyNotes" = @{
            "name" = "Microsoft.MicrosoftStickyNotes"
            "version" = "6.1.2.0"
            "path" = "\\172.20.1.122\share\software\00newpc\40-Microsoft_Store"
            "filename" = "microsoft-sticky-notes-6-1-2-0.msixbundle"
        }
        #�@Photos��OS�����n�D.
        "Photos" = @{
            "name" = "Microsoft.Windows.Photos"
            "version" = "2024.11070.31001.0"
            "path" = "\\172.20.1.122\share\software\00newpc\40-Microsoft_Store"
            "filename" = "microsoft-photos-2024-11070-31001-0.msixbundle"
        }
    }

    foreach ($app in $apps.Keys) {
        Write-Debug "Check appx software (msixbundle): $($apps.$app.name) version: $($apps.$app.version)"
        $installed_version = (Get-AppxPackage -Name $apps.$app.name).Version
        Write-Debug "Installed version: $installed_version"
        $result = Compare-Version -Version1 $apps.$app.version -Version2 $installed_version

        if ($result) {
            Write-Output "��sAppx: $($apps.$app.name)"
            $app_fullpath = "$($apps.$app.path)\$($apps.$app.filename)"
            Write-Debug "App fullpath = $($app_fullpath)"
            Add-AppxPackage -Path $app_fullpath -Update
        } else {
            Write-Output "Appx: $($apps.$app.name) ����s."
        }
    }

}


function remove-apps {
    if ($(Get-OSVersion) -in "Windows 10","Windows 11") {
        Remove-AppsinWin10
    }
}

#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q�פJ�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    
    }

    remove-apps
    update-apps
    pause
}

