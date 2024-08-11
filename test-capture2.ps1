# 加?GDI+?
Add-Type -AssemblyName System.Drawing

# ?取Notepad的窗口句柄
$notepadHandle = (Get-Process notepad).MainWindowHandle

# 定?保存截?的路?
$savePath = "C:\Temp\NotepadScreenshot.png"

# 使用Win32 API?取窗口大小
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

# ?算窗口的?度和高度
$width = $rect.Right - $rect.Left
$height = $rect.Bottom - $rect.Top

# ?建一?位??象?存?截?
$bitmap = New-Object Drawing.Bitmap $width, $height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

# ?屏幕复制窗口的?像
$graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, [System.Drawing.Size]::new($width, $height))

# 保存截??PNG文件
$bitmap.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)

# ?放?象
$bitmap.Dispose()
$graphics.Dispose()

Write-Output "Screenshot saved to $savePath"
