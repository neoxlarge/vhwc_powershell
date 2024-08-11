Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# �w�q���n�� Windows API ���
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class User32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@ -PassThru

# ��� Notepad �i�{
$notepadProcess = Get-Process | Where-Object { $_.ProcessName -eq "notepad" } | Select-Object -First 1

if ($null -eq $notepadProcess) {
    Write-Host "�����B�椤�� Notepad �{�ǡC"
    exit
}

# ��� Notepad ���f�y�`
$handle = $notepadProcess.MainWindowHandle

if ($handle -eq [IntPtr]::Zero) {
    Write-Host "�L�k��� Notepad ���f�y�`�C"
    exit
}

# �N Notepad ���f�s��̤W�h
[User32]::ShowWindow($handle, 9) # 9 �N�� SW_RESTORE
[User32]::SetForegroundWindow($handle)

# �����f�@�Ǯɶ����T��
Start-Sleep -Milliseconds 500

# ����Ҧ��ù����`�j�p
$totalSize = [System.Windows.Forms.SystemInformation]::VirtualScreen

# �Ыؤ@�Ӧ�Ϲ�H
$bitmap = New-Object System.Drawing.Bitmap $totalSize.Width, $totalSize.Height

# �Ыؤ@�ӹϧι�H
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

# �����ӿù�
$graphics.CopyFromScreen($totalSize.Location, [System.Drawing.Point]::Empty, $totalSize.Size)

# �ͦ��ߤ@�����W�]�ϥη�e�ɶ��W�^
$fileName = "c:\temp\FullScreen_With_Notepad_$(Get-Date -Format 'yyyyMMdd_HHmmss').png"

# �O�s�I��
$bitmap.Save($fileName)

# ����귽
$graphics.Dispose()
$bitmap.Dispose()

Write-Host "���ù��I�ϡ]�]�t�e���� Notepad�^�w�O�s��: $fileName"
Write-Host "�I�Ϥj�p: �e=$($totalSize.Width), ��=$($totalSize.Height)"