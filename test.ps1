$shell = New-Object -ComObject Shell.Application
$desktop = $shell.NameSpace("C:\Users\73058\Desktop")
$desktop.Items() | Sort-Object Name | ForEach-Object {
    $desktop.MoveHere($_.Path)
}

