#line_notify_token  = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI"


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



function Send-LineNotifyMessage2 {
    [CmdletBinding()]
    param (
      [Parameter(Mandatory = $true)]
      [string]$Token, # LINE Notify 存取權杖
      [Parameter(Mandatory = $true)]
      [string]$Message, # 要發送的訊息內容
      [string]$ImagePath, # 要傳送的圖片檔案路徑
      [string]$StickerPackageId, # 要一併傳送的貼圖套件 ID
      [string]$StickerId # 要一併傳送的貼圖 ID
    )
  
    # LINE Notify API 的 URI
    $uri = "https://notify-api.line.me/api/notify"
  
    # 建立 HTTP Header
    $headers = New-Object System.Collections.Hashtable
    $headers.Add("Authorization", "Bearer $Token")
    $headers.Add("Content-Type", "multipart/form-data;boundary=$boundary")
  
    # 建立 Form 欄位
    $form = New-Object System.Collections.Hashtable
    $form.Add("message", $Message)
  
    # 如果要傳送圖片
    if ($ImagePath) {
      $fileData = [System.IO.File]::ReadAllBytes($ImagePath)
      $form.Add("imageFile", [PSCustomObject]@{
        "Content-Type" = "image/png"
        "Content-Disposition" = "form-data; name=\'imageFile\'; filename=\'$(Split-Path $ImagePath -Leaf)\'"
        "Value" = $fileData
      })
    }
  
    # 如果要傳送貼圖，加入貼圖套件 ID 和貼圖 ID
    if ($StickerPackageId -and $StickerId) {
      $form.Add("stickerPackageId", $StickerPackageId)
      $form.Add("stickerId", $StickerId)
    }
  
    # 使用 Invoke-WebRequest 傳送 HTTP POST 請求
    Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $form
  }
  

$line_token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI"

#Send-LineNotifyMessage -Token $line_token -Message "vhwc test line "
Send-LineNotifyMessage2 -Token $line_token -Message "haha" #-ImagePath d:\mis\webdriver\vhcy_cpoe_20240311205152.png