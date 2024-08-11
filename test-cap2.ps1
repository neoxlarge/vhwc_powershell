Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 定義必要的 Windows API 函數
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

# 獲取 Notepad 進程
$notepadProcess = Get-Process | Where-Object { $_.ProcessName -eq "notepad" } | Select-Object -First 1

if ($null -eq $notepadProcess) {
    Write-Host "未找到運行中的 Notepad 程序。"
    exit
}

# 獲取 Notepad 窗口句柄
$handle = $notepadProcess.MainWindowHandle

if ($handle -eq [IntPtr]::Zero) {
    Write-Host "無法獲取 Notepad 窗口句柄。"
    exit
}

# 將 Notepad 窗口叫到最上層
[User32]::ShowWindow($handle, 9) # 9 代表 SW_RESTORE
[User32]::SetForegroundWindow($handle)

# 給窗口一些時間來響應
Start-Sleep -Milliseconds 500

# 獲取所有螢幕的總大小
$totalSize = [System.Windows.Forms.SystemInformation]::VirtualScreen

# 創建一個位圖對象
$bitmap = New-Object System.Drawing.Bitmap $totalSize.Width, $totalSize.Height

# 創建一個圖形對象
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

# 捕獲整個螢幕
$graphics.CopyFromScreen($totalSize.Location, [System.Drawing.Point]::Empty, $totalSize.Size)

# 生成唯一的文件名（使用當前時間戳）
$fileName = "c:\temp\FullScreen_With_Notepad_$(Get-Date -Format 'yyyyMMdd_HHmmss').png"

# 保存截圖
$bitmap.Save($fileName)

# 釋放資源
$graphics.Dispose()
$bitmap.Dispose()

Write-Host "全螢幕截圖（包含前景的 Notepad）已保存為: $fileName"
Write-Host "截圖大小: 寬=$($totalSize.Width), 高=$($totalSize.Height)"