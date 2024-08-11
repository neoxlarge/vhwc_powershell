# 加?GDI+?
Add-Type -AssemblyName System.Drawing

# 定?保存截?的路?
$savePath = "C:\Temp\ScreenCapture.png"

# ?取屏幕尺寸
$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height

# ?建一?位??象
$bitmap = New-Object Drawing.Bitmap $screenWidth, $screenHeight

# ?屏幕捕??像
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen(0, 0, 0, 0, [System.Drawing.Size]::new($screenWidth, $screenHeight))

# 保存截?
$bitmap.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)

# ?放?象
$bitmap.Dispose()
$graphics.Dispose()

Write-Output "Screenshot saved to $savePath"
