# �[?GDI+?
Add-Type -AssemblyName System.Drawing

# �w?�O�s�I?����?
$savePath = "C:\Temp\ScreenCapture.png"

# ?���̹��ؤo
$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height

# ?�ؤ@?��??�H
$bitmap = New-Object Drawing.Bitmap $screenWidth, $screenHeight

# ?�̹���??��
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen(0, 0, 0, 0, [System.Drawing.Size]::new($screenWidth, $screenHeight))

# �O�s�I?
$bitmap.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)

# ?��?�H
$bitmap.Dispose()
$graphics.Dispose()

Write-Output "Screenshot saved to $savePath"
