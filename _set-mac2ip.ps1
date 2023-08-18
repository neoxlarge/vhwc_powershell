
param($runadmin)


function Get-MacAddress {
    <#
    只能在172.20.*才能用.
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
        Write-Host "未加入網域!! 請加入網域再執行." -ForegroundColor Red
        pause
        break
    }

    if ($curr_ipconf -eq $null) {
        Write-Host "未發現符合灣橋網路的IP, 請檢查網路設定."
        Pause
        break
    }

    Write-Output "=========================="
    Write-Output "當前電腦ip confing"
    Write-Output "IP: $($curr_ipconf.ip)"
    Write-Output "Mac Address: $($curr_ipconf.mac)"
    Write-Output "=========================="
    $target_ip = Read-Host "請輸入要綁定的IP"

    write-host "查詣DHCP:$dhcp_server 中,請等待..."

    #原本向DHCP查所有的scope, 但是DHCP scope幾乎不會變,所以就抄下來就好, 也減輕DHCP負擔.
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
        # 獲取當前作用域中所有已保留的 IP 地址
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
    Write-Host "IP: $($result.IPAddress) 綁定的MAC Address為"
    write-host "MAC: $($result.ClientID)"
    Write-Host "是否要更改綁定為"
    Write-Host "MAC: $($curr_ipconf.mac)"
    Write-Host "請檢查以上數值是否正確, 下一步將修改DHCP server 上的資料" -ForegroundColor Red
    Write-Host "=========================="
    $yn = Read-Host "請輸入Y/N" 

   

    if ($yn -eq "y") {

        $script_block_setMAC = @{
            ComputerName = $dhcp_server;
            ScriptBlock  = { Set-DhcpServerv4Reservation -IPAddress $args[0] -ClientId $args[1] };
            ArgumentList = @($($result.IPAddress), $($curr_ipconf.mac).replace(":", "-")) #wmi查到的mac用:分隔,改成-號
        }

        Invoke-Command @script_block_setMAC

        if (!$Error) {

            Write-Host "Mac Address綁定完成."

            Write-Host "重新取得IP中,請等待..."

            Invoke-Expression -Command "ipconfig /release"

            Invoke-Expression -Command "ipconfig /renew"
        }

    }



    
}

#檔案獨立執行時會執行函式, 如果是被?入時不會執行函式.
if ($run_main -eq $null) {

    #檢查是否管理員
    $check_admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

    if (!$check_admin -and !$runadmin) {
        #如果非管理員, 就試著run as admin, 並傳入runadmin 參數1. 因為在網域一般使用者永遠拿不是管理員權限, 會造成無限重跑. 此參數用來輔助判斷只跑一次. 
        Start-Process powershell.exe -ArgumentList "-FILE `"$PSCommandPath`" -Executionpolicy bypass -NoProfile  -runadmin 1" -Verb Runas; exit
    }

    set-mac2ip
    pause
}
