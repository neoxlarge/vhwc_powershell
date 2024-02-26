
function uninstall-software {
    <#
    .SYNOPSIS
        移除指定名字的軟體
    
    #>

    [Parameter(Mandatory = $true)]
    [string]$name 


    $mymodule_path = Split-Path $PSCommandPath + "\"
    Import-Module $mymodule_path + "get-installedprogramlist.psm1"
    Import-Module $mymodule_path + "get-admin_cred.psm1"

    $credential = get-admin_cred

    $all_installed_program = get-installedprogramlist
    $software_is_installed = $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like $name }


    if ($software_is_installed -ne $null) {
        $uninstallstring = $software_is_installed.uninstallString.Split(" ")[1].replace("I", "X")

        $running_proc = Start-Process -FilePath "msiexec.exe" -ArgumentList "$uninstallstring /passive" -Credential $credential -PassThru
        $running_proc.WaitForExit()     
    } else {
        return "找不到軟體: $name"
    }

}