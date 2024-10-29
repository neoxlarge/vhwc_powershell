# 有時候smartcard 服務會消失, 把smartcard服務加回去.

function Check-SmartCard {
    $serviceName = "SCardSvr"

    Write-Host "檢查 Smart Card 服務狀態..."

    # 檢查服務是否存在
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if ($null -eq $service) {
        Write-Host "Smart Card 服務不存在。嘗試添加..."
        try {
            Add-WindowsCapability -Online -Name "SmartCard.DiscreteSignalService~~~~0.0.1.0"
            Write-Host "Smart Card 服務已成功添加。"
        }
        catch {
            Write-Host "添加 Smart Card 服務時出錯: $_"
            return
        }
    }

    # 再次獲取服務狀態（以防剛剛添加）
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if ($service.StartType -ne 'Automatic') {
        Write-Host "設置 Smart Card 服務為自動啟動..."
        Set-Service -Name $serviceName -StartupType Automatic
    }

    if ($service.Status -ne 'Running') {
        Write-Host "啟動 Smart Card 服務..."
        Start-Service -Name $serviceName
    }

    # 最終檢查
    $finalStatus = Get-Service -Name $serviceName
    Write-Host "Smart Card 服務當前狀態:"
    Write-Host "  啟動類型: $($finalStatus.StartType)"
    Write-Host "  運行狀態: $($finalStatus.Status)"
}

Check-SmartCard