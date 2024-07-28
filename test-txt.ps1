$computerName = "172.20.2.123"
$community = "public"  # 或者您的 SNMP 社群字串
$oid = "1.3.6.1.2.1.43.10.2.1.4.1.1"  # 這是印表機總頁數的 OID

$result = Get-SnmpData -IP $computerName -OID $oid -Community $community

Write-Host "印表機總頁數: $result"