$localGroup = "Device Owners"
$usernameToAdd = "Domain Users"

# ?取要添加的用??象
$userToAdd = Get-ADGroup -Server wcdc2 -Identity $usernameToAdd

# ?取本地群??象
$group = [ADSI]"WinNT://$env:COMPUTERNAME/$localGroup,group"

$xx = [ADSI]"WinNT://vhcy.gov.tw/Domain Users,group"

# 添加用?到群?
$group.Add($userToAdd.Path)