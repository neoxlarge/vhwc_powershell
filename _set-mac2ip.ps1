
param($runadmin)


function Get-MacAddress {
    <#
    �u��b172.20.*�~���.
    #>

    $mac_addr = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.20.*" } 

    if ($mac_addr) {
    return @{"ip" = $mac_addr.IPAddress[0]; "mac" = $mac_addr.MACaddress }
    } else {return $null}
}


function Set-mac2ip {

    $curr_ipconf = Get-MacAddress
    $dhcp_server = "wcdc2"

    $isJoinAD = (Get-WmiObject -Class Win32_ComputerSystem).PartOfDomain

    if (!$isJoinAD) {
        Write-Host "���[�J����!! �Х[�J����A����." -ForegroundColor Red
        pause
        break
    }

    if ($curr_ipconf -eq $null) {
        Write-Host "���o�{�ŦX�W��������IP, ���ˬd�����]�w."
        Pause
        break
    }

    Write-Output "=========================="
    Write-Output "��e�q��ip confing"
    Write-Output "IP: $($curr_ipconf.ip)"
    Write-Output "Mac Address: $($curr_ipconf.mac)"
    Write-Output "=========================="
    $target_ip = Read-Host "�п�J�n�j�w��IP"

    write-host "�d��DHCP:$dhcp_server ��,�е���..."

    #�쥻�VDHCP�d�Ҧ���scope, ���ODHCP scope�X�G���|��,�ҥH�N�ۤU�ӴN�n, �]�DHCP�t��.
    #$scopes = Invoke-Command -ComputerName $dhcp_server -ScriptBlock { Get-DhcpServerv4Scope }

    $scopes = @("172.20.1.64",
        "172.20.2.0",  
        "172.20.2.64", 
        "172.20.2.128",
        "172.20.2.192",
        "172.20.3.0",  
        "172.20.3.128",
        "172.20.4.0",  
        "172.20.5.0",  
        "172.20.5.128",
        "172.20.7.0",  
        "172.20.8.0",  
        "172.20.9.0",  
        "172.20.11.0", 
        "172.20.12.0", 
        "172.20.13.0", 
        "172.20.15.0", 
        "172.20.16.0", 
        "172.20.17.0", 
        "172.20.18.0", 
        "172.20.19.0", 
        "172.20.20.0", 
        "172.20.34.0", 
        "172.20.35.0", 
        "172.20.64.0", 
        "172.20.65.0", 
        "172.20.66.0") 


    $curr_ip_split = $curr_ipconf.ip.Split(".")[0..2]
    $curr_subnet = "$($curr_ip_split -join ".").*"
    $scopes = $scopes | Where-Object -FilterScript { $_ -like $curr_subnet }

    $script_block = {
        param($scopeId)
        Get-DhcpServerv4Reservation -ScopeId $scopeId
    }

   

    $result = $null
    foreach ($s in $Scopes) {
        # �����e�@�ΰ줤�Ҧ��w�O�d�� IP �a�}
        $ReservedIps = Invoke-Command -ComputerName $dhcp_server -ScriptBlock $script_block -ArgumentList $s
        #Write-Host $s.ScopeId
        
        foreach ($r in $ReservedIps) {

            if ("$($r.IPAddress)" -eq "$target_ip") {
                $result = $r
                #$result | Select-Object -Property * | Write-Host
                break 
               
            }
        }
        if ($result -ne $null) { break }
        
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
            ScriptBlock  = { Set-DhcpServerv4Reservation -IPAddress $args[0] -ClientId $args[1] };
            ArgumentList = @($($result.IPAddress), $($curr_ipconf.mac).replace(":", "-")) #wmi�d�쪺mac��:���j,�令-��
        }

        Invoke-Command @script_block_setMAC

        if (!$Error) {

            Write-Host "Mac Address�j�w����."

            Write-Host "���s���oIP��,�е���..."

            Invoke-Expression -Command "ipconfig /release"

            Invoke-Expression -Command "ipconfig /renew"
        }

    }



    
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
