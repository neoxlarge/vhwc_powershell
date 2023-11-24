#清除DNS快取
ipconfig /flushdns
#ipconfig /renew

#停止所以IE執行
Get-Process -Name iexplore -ErrorAction SilentlyContinue | Stop-Process -Force

#刪掉offline data
Get-ChildItem C:\2100\SSO\OFFLINEDATA | Remove-Item -Recurse -Force

#執行公文環境檔
$pathfile = "\\172.20.5.187\mis\08-2100公文系統\01公文環境檔.exe"
if (Test-Path $pathfile) {
    Copy-Item -Path $pathfile C:\2100 -Force
    & C:\2100\01公文環境檔.exe
}