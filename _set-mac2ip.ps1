
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
    $dhcp_server = "wcdc2"

    Write-Output "=========================="
    Write-Output "��e�q��ip confing"
    Write-Output "IP: $($curr_ipconf.ip)"
    Write-Output "Mac Address: $($curr_ipconf.mac)"
    Write-Output "=========================="
    $target_ip = Read-Host "�п�J�n�j�w��IP"

    write-host "�d��DHCP:$dhcp_server ��,�е���..."

    $Scopes = Invoke-Command -ComputerName $dhcp_server -ScriptBlock { Get-DhcpServerv4Scope }

    $script_block = {
        param($scopeId)
        Get-DhcpServerv4Reservation -ScopeId $scopeId
    }

   

    $result = $null
    foreach ($s in $Scopes) {
        # �����e�@�ΰ줤�Ҧ��w�O�d�� IP �a�}
        $ReservedIps = Invoke-Command -ComputerName $dhcp_server -ScriptBlock $script_block -ArgumentList $s.ScopeId
        #Write-Host $s.ScopeId
        
        foreach ($r in $ReservedIps) {

            if ("$($r.IPAddress)" -eq "$target_ip") {
                $result = $r
                $result | Select-Object -Property * | Write-Host
                
                break 
               
            }
        }
        if ($result -ne $null) {break}
        
    }

    Write-Host "=========================="
    Write-Host "IP: $($result.IPAddress) �j�w��MAC Address��"
    write-host "MAC: $($result.ClientID)"
    Write-Host "�O�_�n���j�w��"
    Write-Host "MAC: $($curr_ipconf.mac)"
    Write-Host "���ˬd�H�W�ƭȬO�_���T, �U�@�B�N�ק�DHCP server �W�����" -ForegroundColor Red
    Write-Host "=========================="
    $yn = Read-Host "�п�JY/N" 

   

    if ($yn -eq "y") {

        $script_block_setMAC = @{
            ComputerName = $dhcp_server;
            ScriptBlock = { Set-DhcpServerv4Reservation -IPAddress $args[0] -ClientId $args[1] };
            ArgumentList = @($($result.IPAddress),$($curr_ipconf.mac).replace(":","-"))
            
        }
        Invoke-Command @script_block_setMAC


        #Set-DhcpServerv4Reservation -ScopeId $result.ScopeId -IPAddress $result.IPAddress -ClientId $curr_ipconf 
        #Invoke-Command -ComputerName $dhcp_serve -ScriptBlock $script_block_setMAC -ArgumentList "-ip '$($result.IPAddress)' -mac '$($curr_ipconf.mac)'"
        #Set-DhcpServerv4Reservation -ScopeId "172.20.5.128" -IPAddress "172.20.5.185" -ClientId "1c-69-7a-3d-50-b3" 

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
