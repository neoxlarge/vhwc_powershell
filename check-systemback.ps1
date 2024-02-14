#機房檢查
#系統備份檢查



function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI",                 # Line Notify 存取權杖

        [Parameter(Mandatory = $true)]
        [string]$Message,               # 要發送的訊息內容

        [string]$StickerPackageId,      # 要一併傳送的貼圖套件 ID

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
        Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payload

        # 訊息成功傳送
        Write-Output "訊息已成功傳送。"
    }
    catch {
        # 發生錯誤，輸出錯誤訊息
        Write-Error $_.Exception.Message
    }
}




$check_list = @("\\172.20.1.122\backup\001-002-hisdb",
                "\\172.20.1.122\backup\001-014-dbSTUDY", 
                "\\172.20.1.122\backup\001-016-homecare",
                "\\172.20.1.122\backup\001-025-nurse",
                "\\172.20.1.122\backup\001-025-pts",
                "\\172.20.1.122\backup\001-067-sk02p",
                "\\172.20.1.122\backup\200-033-hisdb-vghtc"
                )

$path = "\\172.20.1.122\backup\001-002-hisdb"

$today = Get-Date -Format "yyyyMMdd"

$dmp_filename = "exp_full_vhgp_$today.dmp"
                
$latest_file = Get-ChildItem -Path $path | Where-Object -FilterScript {$_.Name -match $dmp_filename}

if ($latest_file) {

    if ($latest_file.Length -gt 40GB) {
        $line_msg = "$latest_file 存在, 檔案大小: $($latest_file.length/(1024*1024*1024))GB."
    } else {
        $line_msg = "$dmp_filename 檔案小於40GB "
    }
    
} else {
    $line_msg = "$dmp_filename 不存在!!!"
}


Send-LineNotifyMessage -Message $line_msg

# \\172.20.1.122\backup\001-014-dbSTUDY\dbSTUDY_Mon.zip