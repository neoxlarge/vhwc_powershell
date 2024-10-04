# 設定要搜索的時間範圍（例如：最近1小時）
$StartTime = (Get-Date).AddHours(-1)

# 從應用程式日誌中獲取錯誤事件
$Events = Get-WinEvent -FilterHashtable @{
    LogName = 'Application'
    Level = 2  # 2 代表錯誤級別
    StartTime = $StartTime
} -ErrorAction SilentlyContinue

# 如果找到錯誤事件，則顯示它們
if ($Events) {
    Write-Host "發現以下應用程式錯誤："
    foreach ($Event in $Events) {
        Write-Host "時間: $($Event.TimeCreated)"
        Write-Host "來源: $($Event.ProviderName)"
        Write-Host "事件 ID: $($Event.Id)"
        Write-Host "消息: $($Event.Message)"
        Write-Host "------------------------"
    }
} else {
    Write-Host "在指定的時間範圍內沒有發現應用程式錯誤。"
}

# 從系統日誌中獲取錯誤事件
$SysEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    Level = 2  # 2 代表錯誤級別
    StartTime = $StartTime
} -ErrorAction SilentlyContinue

# 如果找到系統錯誤事件，則顯示它們
if ($SysEvents) {
    Write-Host "發現以下系統錯誤："
    foreach ($Event in $SysEvents) {
        Write-Host "時間: $($Event.TimeCreated)"
        Write-Host "來源: $($Event.ProviderName)"
        Write-Host "事件 ID: $($Event.Id)"
        Write-Host "消息: $($Event.Message)"
        Write-Host "------------------------"
    }
} else {
    Write-Host "在指定的時間範圍內沒有發現系統錯誤。"
}