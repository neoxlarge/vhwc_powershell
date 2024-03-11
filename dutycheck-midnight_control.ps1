# 建立session到嘉義遠端桌面主機 172.19.1.24
# 在24執行截圖程式存回本地


function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI", # Line Notify 存取權杖

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
        Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payload

        # 訊息成功傳送
        Write-Output "訊息已成功傳送。"
    }
    catch {
        # 發生錯誤，輸出錯誤訊息
        Write-Error $_.Exception.Message
    }
}


$Username = "vhcy\73058"
$Password = "Q1220416-"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

$remote_computer = 'remote_WIN2016.vhcy.gov.tw'
$output_path ='\\172.20.5.185\mis\webdriver'

$script_block = {
    param($output_path)

    write-output $output_path
    #powershell遠端登入後, 不會有\\172.20.5.185\mis的權限, 要掛上磁碟機後才有權限.  
    $Username = "vhcy\73058"
    $Password = "Q1220416-"
    $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

    New-PSDrive -Name Q -Root "$output_path" -Credential $credential -PSProvider FileSystem -Persist
    
    $proc = Start-Process -FilePath "D:\mis\webdriver\dutycheck-midnight_24.exe" -ArgumentList "--output_path $output_path" -PassThru
    $proc.WaitForExit()
    
    Remove-PSDrive -Name "Q"

 }
 
 #Invoke-Command -ComputerName $remote_computer -ScriptBlock $script_block -Credential $credential  -ArgumentList $output_path

$json = Get-Content -Path ($output_path + "\dutycheck.json")
$reprots = ConvertFrom-Json -InputObject $json

foreach ($re in $reprots) {
    $title_message = "$($re.branch) $($re.date) $($re.time)"




    $send_msg = $title_message
    Send-LineNotifyMessage -Message $send_msg
}


