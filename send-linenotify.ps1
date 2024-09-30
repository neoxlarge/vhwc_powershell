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
  
  function Send-LineNotifyMessage3 {
    [CmdletBinding()]
    param (
        [string]$Token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI",
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [string]$StickerPackageId,
        [string]$StickerId,
        [string]$ImagePath
    )

    $uri = "https://notify-api.line.me/api/notify"
    $headers = @{
        "Authorization" = "Bearer $Token"
    }
    
    if ($ImagePath) {
        if (-not (Test-Path $ImagePath)) {
            Write-Error "圖片文件不存在: $ImagePath"
            return
        }

        $fileInfo = Get-Item $ImagePath
        if ($fileInfo.Length -gt 10MB) {
            Write-Error "圖片文件大小超過10MB限制: $($fileInfo.Length) bytes"
            return
        }

        $mimeType = Get-MimeType $ImagePath
        if ($mimeType -notin @("image/jpeg", "image/png")) {
            Write-Error "不支持的圖片格式。只支持 JPEG 和 PNG 格式。"
            return
        }

        $boundary = [System.Guid]::NewGuid().ToString()
        $LF = "`r`n"
        $fileBytes = [System.IO.File]::ReadAllBytes($ImagePath)
        $enc = [System.Text.Encoding]::UTF8
        
        $bodyLines = (
            "--$boundary",
            "Content-Disposition: form-data; name=`"message`"",
            "",
            $Message,
            "--$boundary",
            "Content-Disposition: form-data; name=`"imageFile`"; filename=`"$(Split-Path $ImagePath -Leaf)`"",
            "Content-Type: $mimeType",
            "",
            [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetString($fileBytes),
            "--$boundary--"
        ) -join $LF

        $headers["Content-Type"] = "multipart/form-data; boundary=$boundary"
        $body = $enc.GetBytes($bodyLines)

        Write-Verbose "圖片大小: $($fileBytes.Length) bytes"
        Write-Verbose "圖片 MIME 類型: $mimeType"
    }
    else {
        $body = @{
            "message" = $Message
        }
        if ($StickerPackageId -and $StickerId) {
            $body["stickerPackageId"] = $StickerPackageId
            $body["stickerId"] = $StickerId
        }
    }

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $body -ErrorVariable restError -ErrorAction Stop
        Write-Output "訊息已成功傳送。響應: $($response | ConvertTo-Json -Depth 3)"
    }
    catch {
        Write-Error "發送訊息時發生錯誤: $_"
        if ($restError) {
            Write-Error "REST 錯誤詳情: $($restError.Message)"
        }
        if ($_.Exception.Response) {
            $responseStream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($responseStream)
            $responseBody = $reader.ReadToEnd()
            Write-Error "響應內容: $responseBody"
        }
    }
}

function Get-MimeType($filePath) {
    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
    switch ($extension) {
        ".jpg"  { return "image/jpeg" }
        ".jpeg" { return "image/jpeg" }
        ".png"  { return "image/png" }
        default { return "application/octet-stream" }
    }
}


function Send-LineNotify4 {
  param (
      [string]$token = "",
      [string]$message,
      [string]$imagePath,
      [bool]$notificationDisabled = $true  # 設置通知是否禁用的參數，預設為禁用
  )

  Add-Type -AssemblyName System.Net.Http

  $uri = "https://notify-api.line.me/api/notify"

  # 準備訊息內容
  $body = @{
      message = $message
      notificationDisabled = $notificationDisabled  # 將 notificationDisabled 參數添加到訊息內容中
  }

  # 準備multipart/form-data 格式的內容
  $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
  foreach ($key in $body.Keys) {
      $content = [System.Net.Http.StringContent]::new($body[$key])
      $multipartContent.Add($content, $key)
  }

  # 加入圖片
  if ($imagePath -ne "") {
      $imageStream = [System.IO.File]::OpenRead($imagePath)
      $imageContent = [System.Net.Http.StreamContent]::new($imageStream)
      $imageContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("image/png")  # 可根据实际图片类型调整
      $multipartContent.Add($imageContent, "imageFile", (Split-Path $imagePath -Leaf))
  }

  # 準備HTTP請求
  $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $uri)
  $request.Headers.Authorization = "Bearer $token"
  $request.Content = $multipartContent

  # 發送請求
  $httpClient = [System.Net.Http.HttpClient]::new()
  $response = $httpClient.SendAsync($request).Result

  # 處理回應
  if ($response.IsSuccessStatusCode) {
      Write-Host "訊息發送成功。"
  }
  else {
      Write-Host "無法發送訊息。StatusCode: $($response.StatusCode)，原因: $($response.ReasonPhrase)"
  }

  start-sleep -second 2
}



$line_token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI"

#Send-LineNotifyMessage -Token $line_token -Message "vhwc test line "
Send-LineNotify4 -Token $line_token -Message "haha" -ImagePath C:\temp\FullScreen_With_Notepad_20240811_105006.png