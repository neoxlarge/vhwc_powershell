
param($runadmin)


function Get-MacAddress {
    <#
    �u��b172.*�~���.
    #>

    $mac_addr = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPAddress -ne $null -and $_.IPAddress[0] -like "1*" } 

    return @{"ip" = $mac_addr.IPAddress[0]; "mac" = $mac_addr.MACaddress }
}


function Set-mac2ip {

    $curr_ipconf = Get-MacAddress
    $dhcp_serve = "wcdc2"

    Write-Output "=========================="
    Write-Output "��e�q��ip confing"
    Write-Output "IP: $($curr_ipconf.ip)"
    Write-Output "Mac Address: $($curr_ipconf.mac)"
    Write-Output "=========================="
    $target_ip = Read-Host "�п�J�n�j�w��IP"

    write-host "�d��DHCP:$dhcp_serve ��,�е���..."

    $Scopes = Invoke-Command -ComputerName $dhcp_serve -ScriptBlock { Get-DhcpServerv4Scope }

    $script_block = {
        param($scopeId)
        Get-DhcpServerv4Reservation -ScopeId $scopeId
    }
    
    foreach ($s in $Scopes) {
        # �����e�@�ΰ줤�Ҧ��w�O�d�� IP �a�}
        $ReservedIps = Invoke-Command -ComputerName $dhcp_serve -ScriptBlock $script_block -ArgumentList $s.ScopeId

        foreach ($r in $ReservedIps) {

            if ("$($r.IPAddress)" -eq "$target_ip") {
                $r.ClientId
                break 2
               
            }
        }

    }

    #�VDHCP server���O�d�Ϭd�߹����쪺IP()

    #�ثe��IP
   


}

#�ɮ׿W�߰���ɷ|����禡, �p�G�O�Q?�J�ɤ��|����禡.
if ($run_main -eq $null) {

    #�ˬd�O�_�޲z��
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #�p�G�D�޲z��, �N�յ�run as admin, �öǤJrunadmin �Ѽ�1. �]���b����@��ϥΪ̥û������O�޲z���v��, �|�y���L�����]. ���ѼƥΨӻ��U�P�_�u�]�@��. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    set-mac2ip
    pause
}
