# 導入模組
Import-Module SNMPv3

# 設置連接參數
$ip = "172.20.5.177"
$user = "SNMPv3用戶名"
$authProtocol = "MD5"  # 或 "SHA"
$authPassword = "認證密碼"
$privProtocol = "AES"  # 或 "DES"
$privPassword = "加密密碼"

# 設置 OID（這裡用總頁數作為例子）
$oid = "1.3.6.1.2.1.43.10.2.1.4.1.1"

# 獲取 SNMP 數據
$result = Get-SnmpV3Data -IP $ip -User $user -AuthProtocol $authProtocol -AuthPassword $authPassword -PrivProtocol $privProtocol -PrivPassword $privPassword -OID $oid

# 輸出結果
Write-Host "印表機總頁數: $result"