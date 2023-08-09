#�����n���(Win10)
param($runadmin)

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")

#���oOS������
function Get-OSVersion {
    $os = (Get-WmiObject -Class Win32_OperatingSystem).Caption

    if ($os -like "*Windows 7*") {
        return "Windows 7"
    }
    elseif ($os -like "*Windows 10*") {
        return "Windows 10"
    }
    else {
        return "Unknown OS"
    }
}


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
        "B9ECED6F.ASUSPCAssistant")              

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


function remove-apps {
    if ($(Get-OSVersion) -eq "Windows 10") {
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
    pause
}


<########################################################################################################

$uninstall_list = @{ name = "onedrive"; version = "0" },
#@{ name = "hicos"; version = "3.0.2" },
@{ name = "skype"; version = "0" }

$all_installed_program = get-installedprogramlist


foreach ($i in $uninstall_list) {

    $app = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "*$($i.name)*" }

    if ($app -ne $null) {
    write-output $app.displayname
    $uninstall_string = $app.UninstallString.Split(" ")
    Write-Output $uninstall_string[2]
    Start-Process -FilePath $uninstall_string[0] -ArgumentList $uninstall_string[2] -wait
    }
   }


#>