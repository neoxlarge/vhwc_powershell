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
# firewall裡有些開9080是錯的, 實際上是2788
$idsm_clients = @{
    'wmis-000-pc05' = @{
        'ip'   = '172.20.5.185'
        'port' = '2788' 
        'location' = '資訊室測試'
    }

    'wpha-pha-pc08' = @{
        'ip'   = '172.20.9.78'
        'port' = '2788'
        'location' = '藥劑科 藥袋列印'
        'always_on' = $true
    }

    'wnur-a1w-pc04' = @{
        'ip'   = '172.20.17.14'
        'port' = '2788'
        'location' = 'A1'
        'always_on' = $true
    }

    'wnur-a2w-pc04' = @{
        'ip'   = '172.20.17.24'
        'port' = '2788'
        'location' = 'A2'
        'always_on' = $true
    }

    'wnur-a3w-pc05' = @{
        'ip'   = '172.20.17.35'
        'port' = '2788'
        'location' = 'A3'
        'always_on' = $true
    }

    'wnur-a5w-pc02' = @{
        'ip'   = '172.20.17.52'
        'port' = '2788'
        'location' = 'A5'
        'always_on' = $true
    }

    'wnur-b1w-pc04' = @{
        'ip'   = '172.20.2.93'
        'port' = '2788'
        'location' = 'B1'
        'always_on' = $true
    }

    'wnur-b2w-pc05' = @{
        'ip'   = '172.20.2.94'
        'port' = '2788'
        'location' = 'B2'
        'always_on' = $true
    }

    'wnur-b3w-pc04' = @{
        'ip'   = '172.20.2.97'
        'port' = '2788'
        'location' = 'B3'
        'always_on' = $true
    }

    'wnur-b5w-pc05' = @{
        'ip'   = '172.20.2.77'
        'port' = '2788'
        'location' = 'B5'
        'always_on' = $true
    }

    'wadm-mrr-pc02' = @{
        'ip'   = '172.20.2.207'
        'port' = '2788'
        'location' = '病歷室'
    }

    'wnur-erx-pc02' = @{
        'ip'   = '172.20.3.3'
        'port' = '2788'
        'location' = '急診室'
        'always_on' = $true
    }

    'wnur-icu-pc01' = @{
        'ip'   = '172.20.5.2'
        'port' = '2788'
        'location' = 'ICU'
        'always_on' = $true
    }

    'wnur-m3w-pc05' = @{
        'ip'   = '172.20.5.31'
        'port' = '2788'
        'location' = 'M3'
        'always_on' = $true
    }

    'wnur-m5a-pc06' = @{
        'ip'   = '172.20.5.26'
        'port' = '2788'
        'location' = 'M5A'
        'always_on' = $true
    }

    'wnur-m5b-pc06' = @{
        'ip'   = '172.20.5.15'
        'port' = '2788'
        'location' = 'M5B'
        'always_on' = $true
    }

    'wlab-000-pc03' = @{
        'ip'   = '172.20.7.18'
        'port' = '2788'
        'location' = '檢驗科血庫'
        'always_on' = $true
    }

    'wlab-000-pc06a' = @{
        'ip'   = '172.20.3.211'
        'port' = '2788'
        'location' = '檢驗科'
        'always_on' = $true
    }

    <# 儀器連線用電腦, 目前應該沒有用到分散式列印, 暫時拿掉
    'wlab-000-pc09' = @{
        'ip'   = '172.20.3.149'
        'port' = '2788'
        'location' = '檢驗科 收信儀器連線'
    }
    #>
    
    'wadm-reg-pc04' = @{
        'ip'   = '172.20.3.29'
        'port' = '2788'
        'location' = '掛號室 代收嘉榮費用'
    }

    'wpsy-pnp-pc01' = @{
        'ip'   = '172.20.3.158'
        'port' = '2788'
        'location' = '精神科'
    }

    'wadm-reg-pc01' = @{
        'ip'   = '172.20.3.77'
        'port' = '2788'
        'location' = '掛號室 曉婷'
    }
    
    'wadm-reg-pc02' = @{
        'ip'   = '172.20.3.1'
        'port' = '2788'
        'location' = '掛號室 玉尊'
    }

    'wadm-reg-pc03' = @{
        'ip'   = '172.20.3.13'
        'port' = '2788'
        'location' = '掛號室 舒璇'
    }

    'wadm-reg-pc04xxxx' = @{
        'ip'   = '172.20.3.120'
        'port' = '2788'
        'location' = '掛號室 第4櫃台'
    }

    'wpha-pha-pc06' = @{
        'ip'   = '172.20.9.76'
        'port' = '2788'
        'location' = '藥劑科 藥局發藥櫃台'
        'always_on' = $true
    }

    'wreh-000-pc03' = @{
        'ip'   = '172.20.17.63'
        'port' = '2788'
        'location' = '復建科 櫃台'
    }
    
    'wdie-out-pc01' = @{
        'ip'   = '172.20.2.138'
        'port' = '2788'
        'location' = '營養室外包商'
        'always_on' = $true
    }
    
    'wreh-000-pc02' = @{
        'ip'   = '172.20.2.150'
        'port' = '2788'
        'location' = '營養室 餐卡, 出餐單'
        'always_on' = $true
    }
}    

function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "ZAxQqfCDIuTL7MzURX1pKTuciEOqnwMqy8lnHNJXEMF", # Line Notify 存取權杖

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
        $location = $clientInfo.location
        $always_on = $clientInfo.always_on

        # 檢查IP是否可Ping通
        if (Test-Connection -ComputerName $ipAddress -Count 2 -Quiet) {
            Write-Debug "Ping $clientName ($ipAddress) 正常."
            
            # 檢查Port服務是否有回應
            $connectionTestResult = Test-NetConnection -ComputerName $ipAddress -Port $portNumber 
            if ($connectionTestResult.TcpTestSucceeded) {
                Write-Debug "  - Port $portNumber 回應正常."
            }
            else {
                Write-Debug "  - Port $portNumber 無回應.發送 Line 通知"
                
                # 發送 Line 通知
                $msg = "🚨分散式client `nName: $clientName `n"
                $msg += "IP Status: $($ipaddress) ping正常 `n"
                if ($always_on -eq $true) {
                    $msg += "Port Status: $portNumber 無回應, 注意此機須在線! `n"
                } else {
                    $msg += "Port Status: $portNumber 無回應 `n" 
                }
                $msg += "Location: $location"

                Send-LineNotifyMessage -Token $line_token -Message $msg
            }
        }
        else {
            Write-Debug "$clientName ($ipAddress) Ping不到,發送 Line 通知 "
       
            # 發送 Line 通知
            $msg = "🚨分散式client `nName: $clientName `n"
            if ($always_on -eq $true) {
                $msg += "IP Status: $($ipaddress) Ping不到, 注意此機須在線! `n"
            } else {
                $msg += "IP Status: $($ipaddress) Ping不到 `n"
            }
            $msg += "Location: $location"

            Send-LineNotifyMessage -Token $line_token -Message $msg
        }
    }
}

schedulecheck-idmsclients -idms_clients $idsm_clients



