
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

# ?建窗体
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "更改DHCP保留IP的MAC地址"
$Form.Size = New-Object System.Drawing.Size(400, 200)
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

# ?建IP地址??和文本框
$IPLabel = New-Object System.Windows.Forms.Label
$IPLabel.Text = "IP地址:"
$IPLabel.Location = New-Object System.Drawing.Point(30, 30)
$Form.Controls.Add($IPLabel)

$IPTextBox = New-Object System.Windows.Forms.TextBox
$IPTextBox.Location = New-Object System.Drawing.Point(120, 30)
$Form.Controls.Add($IPTextBox)

# ?建MAC地址??和文本框
$MACLabel = New-Object System.Windows.Forms.Label
$MACLabel.Text = "MAC地址:"
$MACLabel.Location = New-Object System.Drawing.Point(30, 60)
$Form.Controls.Add($MACLabel)

$MACTextBox = New-Object System.Windows.Forms.TextBox
$MACTextBox.Location = New-Object System.Drawing.Point(120, 60)
$Form.Controls.Add($MACTextBox)

# ?建确定按?
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Text = "确定"
$OKButton.Location = New-Object System.Drawing.Point(150, 100)
$OKButton.Add_Click({
    $IP = $IPTextBox.Text
    $MAC = $MACTextBox.Text

    # ?行更改保留IP的MAC地址的命令
    Set-DhcpServerv4Reservation -IPAddress $IP -ClientId $MAC

    # ?示成功消息框
    [System.Windows.Forms.MessageBox]::Show("已成功更改保留IP的MAC地址。")
})
$Form.Controls.Add($OKButton)

# ?示窗体
$Form.ShowDialog()
