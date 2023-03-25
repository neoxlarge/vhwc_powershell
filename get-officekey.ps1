$OfficeKey = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Office\11.0\Registration\*' | ForEach-Object {Get-ItemProperty $_.PSPath} | Where-Object {$_.DigitalProductId -ne $null} | Select-Object DigitalProductId

$ProductKey = ""

ForEach ($char in $OfficeKey.DigitalProductId) {
    $ProductKey = $ProductKey + [System.String]::Format("{0:X2}", $char)
}

"Office 2003 安裝序號: $ProductKey"