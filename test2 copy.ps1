
function Send-LineNotifyMessage {
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

    # 准备 HTTP 请求
    $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $uri)
    $request.Headers.Authorization = "Bearer $token"

    # 添加消息
    $formData = [System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[string,string]]]::new()
    $formData.Add([System.Collections.Generic.KeyValuePair[string,string]]::new("message", $message))
    $request.Content = [System.Net.Http.FormUrlEncodedContent]::new($formData)

    # 发送请求
    $httpClient = [System.Net.Http.HttpClient]::new()
    $response = $httpClient.SendAsync($request).Result

    # 处理响应
    if (-not $response.IsSuccessStatusCode) {
        Write-Host "Failed to send message. StatusCode: $($response.StatusCode), Reason: $($response.ReasonPhrase)"
        return
    }

    # 如果指定了图片路径，则发送图片
    if ($imagePath -ne "") {
        $imageBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($imagePath))
        $formDataImage = [System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[string,string]]]::new()
        $formDataImage.Add([System.Collections.Generic.KeyValuePair[string,string]]::new("imageFile", $imageBase64))

        $requestImage = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $uri)
        $requestImage.Headers.Authorization = "Bearer $token"
        $requestImage.Content = [System.Net.Http.FormUrlEncodedContent]::new($formDataImage)

        $responseImage = $httpClient.SendAsync($requestImage).Result

        if (-not $responseImage.IsSuccessStatusCode) {
            Write-Host "Failed to send image. StatusCode: $($responseImage.StatusCode), Reason: $($responseImage.ReasonPhrase)"
            return
        }
    }

    Write-Host "Message sent successfully."
}
# 使用示例
$token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI"
$message = "Hello, World!"
$imagePath = "D:\mis\webdriver\vhcy_cpoe_20240311205152.png"
Send-LineNotifyMessage -token $token -message $message -imagePath $imagePath


#Send-LineNotifyMessage -Message "helloxxx" #-ImagePath D:\mis\webdriver\abc.jpegdddd####