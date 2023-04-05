function Get-IPv4Address {
    <#
    只能在172.*才能用.
    #>

    $ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
          Where-Object {$_.IPAddress -ne $null -and $_.IPAddress[0] -like "172.*" } |
          Select-Object -ExpandProperty IPAddress |
          Where-Object { $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}" } |
          Select-Object -First 1
    return $ip
}

write-host (Get-IPv4Address)