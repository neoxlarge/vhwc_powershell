
param($runadmin)
write-host($PSCommandPath)
Import-Module ("\\172.20.5.185\powershell\vhwc_powershell\get-installedprogramlist.psm1")

function uninstall-office2003 {

    # uninstall 2007 office system ?�容?��?�?

    $Username = "vhcy\vhwcmis"
    $Password = "Mis20190610"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    # uninstall 2007 office system ?�容?��?�?
    $software_name = "2007 Office System*"

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -ne $null) {
        $uninstallstring = $software_is_installed.uninstallString.Split(" ")[1]

        $running_proc = Start-Process -FilePath "msiexec.exe" -ArgumentList "$uninstallstring /passive" -Credential $credential -PassThru
        $running_proc.WaitForExit()     
    }


    # uninstall office 2003
    $software_name = "Microsoft Office Professional Edition 2003*"

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $software_name }

    if ($software_is_installed -ne $null) {
        $uninstallstring = $software_is_installed.uninstallString.Split(" ")[1]

        $running_proc = Start-Process -FilePath "msiexec.exe" -ArgumentList "$uninstallstring /passive" -Credential $credential -PassThru
        $running_proc.WaitForExit()     
    }


}

#檔�??��??��??��??��??��?, 如�??�被?�入?��??�執行函�?
if ($run_main -eq $null) {

    #檢查?�否管�???
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如�??�管?�員, 就試?�run as admin, 並傳?�runadmin ?�數1. ?�為?�網?��??�使?�者永?�拿不是管�??��??? ?�造�??��??��?. 此�??�用來�??�判?�只跑�?�? 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    uninstall-office2003
    pause
}
