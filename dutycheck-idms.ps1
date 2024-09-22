# 檢查分散式列印的client端是否有開機及執行
# 1. 用ping檢查client端是否有開機
# 2. 用port 9100檢查client端是否有執行

# 互斥鎖, 為了避免程式重覆執行, 只能有一個執行在執行中.
$mutexName = "Global\dutycheck-idms"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)
if ($mutex.WaitOne(0, $false) -eq $false) { 
    Write-Host "分散式列印檢查通知己在執行中,結束."
    start-sleep -Seconds 5
    exit 
}


# 設定值
# debug log的資訊, Continue會顯示, SilentContinue不會顯示.
$DebugPreference = "Continue"
# 設定Line Notify的Token
$line_token = "ZAxQqfCDIuTL7MzURX1pKTuciEOqnwMqy8lnHNJXEMF"
# 定時的時間
$timer_hours = @(8..17) #8點到17點
$timer_minutes = @(0, 15, 30, 45)



# 分散式列印的客戶端
$idsm_clients = @{
    "wmis-000-pc05" = @{
        "ip"   = "172.20.5.185"
        "port" = "2788" # or port 9080
    }

    "wnur-b3w-pc04" = @{
        "ip"   = "172.20.2.97"
        "port" = "2788"
    }
}    

function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz", # Line Notify 存取權杖

        [Parameter(Mandatory = $true)]
        [string]$Message, # 要發送的訊息內容

        [string]$StickerPackageId, # 要一併傳送的貼圖套件 ID

        [string]$StickerId              # 要一併傳送的貼圖 ID
    )

    # Line Notify API 的 URI
    $uri = "https://notify-api.line.me/api/notify"

    # 設定 HTTP Header，包含 Line Notify 存取權杖
    $headers = @{ "Authorization" = "Bearer $Token" }

    # 設定要傳送的訊息內容
    $payload = @{
        "message" = $Message
    }

    # 如果要傳送貼圖，加入貼圖套件 ID 和貼圖 ID
    if ($StickerPackageId -and $StickerId) {
        $payload["stickerPackageId"] = $StickerPackageId
        $payload["stickerId"] = $StickerId
    }

    try {
        # 使用 Invoke-RestMethod 傳送 HTTP POST 請求
        $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payload
        
        # 訊息成功傳送
        Write-Debug "訊息已成功傳送。"
    }
    catch {
        # 發生錯誤，輸出錯誤訊息
        Write-Error $_.Exception.Message
    }
}


function schedulecheck-idmsclients {

    param (
        $idms_clients
    )

    # 迴圈檢查每個客戶端
    foreach ($clientName in $idms_clients.Keys) {
        $clientInfo = $idsm_clients[$clientName]
        $ipAddress = $clientInfo.ip
        $portNumber = $clientInfo.port

        # 檢查IP是否可Ping通
        if (Test-Connection -ComputerName $ipAddress -Count 2 -Quiet) {
            Write-Host "$clientName ($ipAddress) is reachable."
            $idsm_clients[$clientName]["reachable"] = $true # 將結果寫入 $idsm_clients

            # 檢查Port服務是否有回應
            $connectionTestResult = Test-NetConnection -ComputerName $ipAddress -Port $portNumber
            if ($connectionTestResult.TcpTestSucceeded) {
                Write-Host "  - Port $portNumber is open and responding."
                $idsm_clients[$clientName]["portResponding"] = $true # 將結果寫入 $idsm_clients
            }
            else {
                Write-Host "  - Port $portNumber is not responding."
                $idsm_clients[$clientName]["portResponding"] = $false # 將結果寫入 $idsm_clients
            }
        }
        else {
            Write-Host "$clientName ($ipAddress) is not reachable."
            $idsm_clients[$clientName]["reachable"] = $false # 將結果寫入 $idsm_clients
            $idsm_clients[$clientName]["portResponding"] = $false # 將結果寫入 $idsm_clients
        }
    }




}





