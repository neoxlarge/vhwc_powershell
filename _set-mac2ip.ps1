
param($runadmin)


function Get-MacAddress {
    <#
    只能在172.*才能用.
    #>

    $mac_addr = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
    Where-Object { $_.IPAddress -ne $null -and $_.IPAddress[0] -like "1*" } 

    return @{"ip" = $mac_addr.IPAddress[0]; "mac" = $mac_addr.MACaddress }
}


function Set-mac2ip {

    $curr_ipconf = Get-MacAddress
    $dhcp_serve = "wcdc2"

    Write-Output "=========================="
    Write-Output "當前電腦ip confing"
    Write-Output "IP: $($curr_ipconf.ip)"
    Write-Output "Mac Address: $($curr_ipconf.mac)"
    Write-Output "=========================="
    $target_ip = Read-Host "請輸入要綁定的IP"

    write-host "查詣DHCP:$dhcp_serve 中,請等待..."

    $Scopes = Invoke-Command -ComputerName $dhcp_serve -ScriptBlock { Get-DhcpServerv4Scope }

    $script_block = {
        param($scopeId)
        Get-DhcpServerv4Reservation -ScopeId $scopeId
    }
    
    foreach ($s in $Scopes) {
        # 獲取當前作用域中所有已保留的 IP 地址
        $ReservedIps = Invoke-Command -ComputerName $dhcp_serve -ScriptBlock $script_block -ArgumentList $s.ScopeId

        foreach ($r in $ReservedIps) {

            if ("$($r.IPAddress)" -eq "$target_ip") {
                $result = $r
                $r | gm
                break 2
               
            }
        }
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

        #Set-DhcpServerv4Reservation -ScopeId "172.20.5.0" -IPAddress "保留 IP 地址" -ClientId "ClientID" -MacAddress "新的 MAC 地址"



    }



    
    

    



    #向DHCP server的保留區查詢對應到的IP()

    #目前的IP
   


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
