# 搜尋 DHCP 伺服器上已保留的 IP 地址
# 通過 MAC 地址進行搜尋
# 此ps1檔只能在dhcp server上正確執行. 因為一般電腦上沒有安裝dhcp相關cmdlet.

param($macaddress, $dhcpserver = "wcdc2.vhcy.gov.tw") #可傳入的參數


Function Find-DhcpReservedIPByMacAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MacAddress, # 要搜尋的 MAC 地址
        [string]$DHCPserver  # 指定的server
    )

    # 獲取 DHCP 服務器上所有作用域
    $Scopes = Get-DhcpServerv4Scope -ComputerName $DHCPserver | Select-Object -ExpandProperty ScopeId

    # 遍歷所有作用域
    foreach ($Scope in $Scopes) {
        # 獲取當前作用域中所有已保留的 IP 地址
        $ReservedIps = Get-DhcpServerv4Reservation -ComputerName $DHCPserver -ScopeId $Scope | Select-Object -Property ClientId,IPAddress

        # 遍歷所有已保留的 IP 地址
        foreach ($ReservedIp in $ReservedIps) {
            # 將 MAC 地址轉換為统一格式（不带“-”且全改成小寫字母）
            $Mac = $ReservedIp.ClientId.Replace("-", "").ToLower()
            $MacAddress = $MacAddress.Replace("-", "").ToLower()

            # 如果找到了匹配的 MAC 地址，則返回對應的 IP 地址和作用域
            if ($Mac -eq $MacAddress) {
                Write-Output "作用域： $Scope, 保留 IP 地址： $($ReservedIp.IPAddress)"
                return @{
                    Scope = $Scope
                    IPAddress = $ReservedIp.IPAddress
                }
            }
        }
    }

    # 如果未找到保留 IP 地址，則输出错誤訊息
    Write-Error "未找到 MAC 地址為 $MacAddress 的保留 IP 地址"
}


find-dhcpreservedipbymacaddress -MacAddress $macaddress -DHCPserver $dhcpserver