
function Send-LineNotify {
    param (
        [string]$token,
        [string]$message,
        [string]$imagePath
    )

    $uri = "https://notify-api.line.me/api/notify"

    # 准备消息内容
    $body = @{
        message = $message
    }

    # 准备 multipart/form-data 格式的消息体
    $multipartContent = [System.Net.Http.MultipartFormDataContent]::new()
    foreach ($key in $body.Keys) {
        $content = [System.Net.Http.StringContent]::new($body[$key])
        $multipartContent.Add($content, $key)
    }

    # 添加图片文件
    if ($imagePath -ne "") {
        $imageStream = [System.IO.File]::OpenRead($imagePath)
        $imageContent = [System.Net.Http.StreamContent]::new($imageStream)
        $imageContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::new("image/png")  # 可根据实际图片类型调整
        $multipartContent.Add($imageContent, "imageFile", (Split-Path $imagePath -Leaf))
    }

    # 准备 HTTP 请求
    $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $uri)
    $request.Headers.Authorization = "Bearer $token"
    $request.Content = $multipartContent

    # 发送请求
    $httpClient = [System.Net.Http.HttpClient]::new()
    $response = $httpClient.SendAsync($request).Result

    # 处理响应
    if ($response.IsSuccessStatusCode) {
        Write-Host "Message sent successfully."
    } else {
        Write-Host "Failed to send message. StatusCode: $($response.StatusCode), Reason: $($response.ReasonPhrase)"
    }
}

# 使用示例
$token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI"
$message = "Hello, World!"
$imagePath = "D:\mis\webdriver\vhcy_cpoe_20240311205152_3.png"
Send-LineNotifyMessage -token $token -message $message -imagePath $imagePath


#Send-LineNotifyMessage -Message "helloxxx" #-ImagePath D:\mis\webdriver\abc.jpegdddd####