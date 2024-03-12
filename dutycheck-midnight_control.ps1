# 建立session到嘉義遠端桌面主機 172.19.1.24
# 在24執行截圖程式存回本地
#AVt3SxMcHhatY2fuG2j6HzKGdb5BOTmrfAlEiBolQOO
#CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI
function Send-LineNotify {
    param (
        [string]$token = "AVt3SxMcHhatY2fuG2j6HzKGdb5BOTmrfAlEiBolQOO",
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
    }
    else {
        Write-Host "Failed to send message. StatusCode: $($response.StatusCode), Reason: $($response.ReasonPhrase)"
    }

    start-sleep -second 2
}


$Username = "vhcy\73058"
$Password = "Q1220416-"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

$remote_computer = 'remote_WIN2016.vhcy.gov.tw'
$output_path = '\\172.20.5.185\mis\webdriver'

$script_block = {
    param($output_path)

    write-output $output_path
    #powershell遠端登入後, 不會有\\172.20.5.185\mis的權限, 要掛上磁碟機後才有權限.  
    $Username = "vhcy\73058"
    $Password = "Q1220416-"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    New-PSDrive -Name Q -Root "$output_path" -Credential $credential -PSProvider FileSystem -Persist
    
    $proc = Start-Process -FilePath "D:\mis\webdriver\dutycheck-midnight_24.exe" -ArgumentList "--output_path $output_path\" -PassThru
    $proc.WaitForExit()
    
    Remove-PSDrive -Name "Q"

}
 
Invoke-Command -ComputerName $remote_computer -ScriptBlock $script_block -Credential $credential  -ArgumentList $output_path

$json = Get-Content -Path "$output_path\dutycheck.json"
$reprots = ConvertFrom-Json -InputObject $json

foreach ($re in $reprots) {
    $msg = "$($re.branch)  $($re.checkitem)`n ===$($re.date) $($re.time)=== `n"

    if ($re.result -eq $true) {
        $msg += "🟢 Pass: "
    }
    else {
        $msg += "🚨 Fail: $($re.message)"
    }

    $send_msg = $msg
    Send-LineNotify -Message $send_msg
    if ($re.crop_images.count -ne 0) {
        foreach ($png in $re.crop_images) {
            $msg = "$($re.crop_images.IndexOf($png) + 1)/$($re.crop_images.Count)"
            Send-LineNotify -Message $msg -ImagePath $png
        }
    } else {
        Send-LineNotify -Message $re.branch -ImagePath $re.png_filepath
    }


}


