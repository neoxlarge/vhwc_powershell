
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

Add-Type -AssemblyName System.Windows.Forms

# ?�ص��^
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "���DHCP�O�dIP��MAC�a�}"
$Form.Size = New-Object System.Drawing.Size(400, 200)
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# ?��IP�a�}??�M�奻��
$IPLabel = New-Object System.Windows.Forms.Label
$IPLabel.Text = "IP�a�}:"
$IPLabel.Location = New-Object System.Drawing.Point(30, 30)
$Form.Controls.Add($IPLabel)

$IPTextBox = New-Object System.Windows.Forms.TextBox
$IPTextBox.Location = New-Object System.Drawing.Point(120, 30)
$Form.Controls.Add($IPTextBox)

# ?��MAC�a�}??�M�奻��
$MACLabel = New-Object System.Windows.Forms.Label
$MACLabel.Text = "MAC�a�}:"
$MACLabel.Location = New-Object System.Drawing.Point(30, 60)
$Form.Controls.Add($MACLabel)

$MACTextBox = New-Object System.Windows.Forms.TextBox
$MACTextBox.Location = New-Object System.Drawing.Point(120, 60)
$Form.Controls.Add($MACTextBox)

# ?���̩w��?
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Text = "�̩w"
$OKButton.Location = New-Object System.Drawing.Point(150, 100)
$OKButton.Add_Click({
    $IP = $IPTextBox.Text
    $MAC = $MACTextBox.Text

    # ?����O�dIP��MAC�a�}���R�O
    Set-DhcpServerv4Reservation -IPAddress $IP -ClientId $MAC

    # ?�ܦ��\������
    [System.Windows.Forms.MessageBox]::Show("�w���\���O�dIP��MAC�a�}�C")
})
$Form.Controls.Add($OKButton)

# ?�ܵ��^
$Form.ShowDialog()
