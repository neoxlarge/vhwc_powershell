# 建立session到嘉義遠端桌面主機 172.19.1.24
# 在24執行截圖程式存回本地
# line token(灣橋檢查群組): HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz
# line token(測試1): CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI
# line token(測試2): AVt3SxMcHhatY2fuG2j6HzKGdb5BOTmrfAlEiBolQOO
# 定時每天晚上11:20分, 和早上0點20分執行.


function Send-LineNotify {
    param (
        [string]$token = "HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz",
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



$Username = "vhcy\73058"
$Password = "Q1220416-"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

$remote_computer = 'remote_WIN2016.vhcy.gov.tw'
$output_path = '\\172.20.5.185\mis\webdriver'

#遠端電腦上要執行的指令
$script_block = {
    param($output_path)

    write-output $output_path
    #powershell遠端登入後, 不會有\\172.20.5.185\mis的權限, 要掛上磁碟機後才有權限.  
    $Username = "vhcy\73058"
    $Password = "Q1220416-"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
    New-PSDrive -Name Q -Root "$output_path" -Credential $credential -PSProvider FileSystem 

    # 比對檔案是否一致, 以5.185上的為準.
    $file1 = "\\172.20.5.185\powershell\vhwc_powershell\dist\dutycheck-midnight_24.exe"
    $file2 = "D:\mis\webdriver\dutycheck-midnight_24.exe"

    $file1hash = Get-FileHash -Path $file1
    $file2hash = Get-FileHash -Path $file2

    if ($file1hash.hash -ne $file2hash.Hash) {
        Copy-Item -Path $file1 -Destination  $file2 -Force -Verbose 
    }

    #執行截圖程式
    $proc = Start-Process -FilePath "D:\mis\webdriver\dutycheck-midnight_24.exe" -ArgumentList "--output_path $output_path\" -PassThru
    $proc.WaitForExit()
    
    Remove-PSDrive -Name "Q"

}

# 對遠端電腦丟出要執行的指令區塊
Invoke-Command -ComputerName $remote_computer -ScriptBlock $script_block -Credential $credential  -ArgumentList $output_path

# 讀取截圖報告
$json = Get-Content -Path "$output_path\dutycheck.json"
$reprots = ConvertFrom-Json -InputObject $json

# 依內容發出line notify
foreach ($re in $reprots) {
    $msg = "$($re.branch)  $($re.checkitem)`n ===$($re.date) $($re.time)=== `n"

    if ($re.result -eq $true) {
        $msg += "🟢 Pass: "
    }
    else {
        $msg += "🚨 Fail: $($re.message)"
    }

    $send_msg = $msg
    Send-LineNotify -Message $send_msg -notificationDisabled $false # 發結果時不要靜音, 底下發圖片時靜音,才不會太吵.
    
    if ($re.crop_images.count -ne 0) {
        foreach ($png in $re.crop_images) {
            $msg = "$($re.crop_images.IndexOf($png) + 1)/$($re.crop_images.Count)"
            Send-LineNotify -Message $msg -ImagePath $png
        }
    } else {
        Send-LineNotify -Message $re.branch -ImagePath $re.png_filepath
    }


}


