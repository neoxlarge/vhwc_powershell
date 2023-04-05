#≤æ∞£≥n≈È•Œ.

Import-Module ((Split-Path $PSCommandPath) + "\get-installedprogramlist.psm1")



$uninstall_list = @{ name = "onedrive"; version = "0" },
@{ name = "hicos"; version = "3.0.2" }

$all_installed_program = get-installedprogramlist


foreach ($i in $uninstall_list) {

    $all_installed_program | Where-Object -FilterScript { $_.DisplayName -like "*$($i.name)*" }

 
}
