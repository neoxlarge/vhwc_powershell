
function Send-WoLPacket {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MacAddress,
        [Parameter(Mandatory = $false)]
        [int]$Port = 9
    )

    $macBytes = $MacAddress -split '[:-]' | ForEach-Object { [byte]('0x' + $_) }

    $udpClient = New-Object System.Net.Sockets.UdpClient
    $udpClient.Connect(([System.Net.IPAddress]::Broadcast), $Port)

    $magicPacket = [byte[]](,0xFF * 6 + $macBytes * 16)
    $udpClient.Send($magicPacket, $magicPacket.Length)

    $udpClient.Close()
}

function send-m {

    $Mac = "C0:3F:B5:54:B0:62"
$MacByteArray = $Mac -split "[:-]" | ForEach-Object { [Byte] "0x$_"}
[Byte[]] $MagicPacket = (,0xFF * 6) + ($MacByteArray  * 16)
$UdpClient = New-Object System.Net.Sockets.UdpClient
$UdpClient.Connect(([System.Net.IPAddress]::Broadcast),7)
$UdpClient.Send($MagicPacket,$MagicPacket.Length)
$UdpClient.Close()
}