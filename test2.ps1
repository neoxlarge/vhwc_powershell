$localGroup = "Device Owners"
$usernameToAdd = "Domain Users"

# ?���n�K�[����??�H
$userToAdd = Get-ADGroup -Server wcdc2 -Identity $usernameToAdd

# ?�����a�s??�H
$group = [ADSI]"WinNT://$env:COMPUTERNAME/$localGroup,group"

$xx = [ADSI]"WinNT://vhcy.gov.tw/Domain Users,group"

# �K�[��?��s?
$group.Add($userToAdd.Path)

$acc = Get-Content .\admin.json -Raw
$js = ConvertFrom-Json -InputObject $acc