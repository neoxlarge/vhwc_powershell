# �[?GDI+?
Add-Type -AssemblyName System.Drawing

# ?��Notepad�����f�y�`
$notepadHandle = (Get-Process notepad).MainWindowHandle

# �w?�O�s�I?����?
$savePath = "C:\Temp\NotepadScreenshot.png"

# �ϥ�Win32 API?�����f�j�p
$rect = New-Object RECT
$GetWindowRect = Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class User32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    }
    public struct RECT {
        public int Left, Top, Right, Bottom;
    }
"@ -PassThru

$GetWindowRect::GetWindowRect($notepadHandle, [ref]$rect) | Out-Null

# ?�ⵡ�f��?�שM����
$width = $rect.Right - $rect.Left
$height = $rect.Bottom - $rect.Top

# ?�ؤ@?��??�H?�s?�I?
$bitmap = New-Object Drawing.Bitmap $width, $height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

# ?�̹��`��f��?��
$graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, [System.Drawing.Size]::new($width, $height))

# �O�s�I??PNG���
$bitmap.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)

# ?��?�H
$bitmap.Dispose()
$graphics.Dispose()

Write-Output "Screenshot saved to $savePath"
