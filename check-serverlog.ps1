$today = get-date

$serverlog_checklist = [ordered]@{
    "200-033-hisdb-error"           = @{
        "root_path" = "\\172.20.1.122\log\server-log\200-033-hisdb\"
        "date_path" = "$($today.tostring('yyyy'))\$($today.ToString('yyyy-MM'))\"
        "file_name" = "$($today.ToString('yyyy-MM'))-error.log"
    }

    "200-033-hisdb-ALL-error"       = @{
        "root_path" = "\\172.20.1.122\log\server-log\200-033-hisdb\"
        "date_path" = "$($today.tostring('yyyy'))\$($today.ToString('yyyy-MM'))\"
        "file_name" = "$($today.ToString('yyyy-MM'))-ALL-error.log"
    }

    "200-033-hisdb-vghtc-error"     = @{
        "root_path" = "\\172.20.1.122\log\server-log\200-033-hisdb-vghtc\"
        "date_path" = "$($today.tostring('yyyy'))\$($today.ToString('yyyy-MM'))\"
        "file_name" = "$($today.ToString('yyyy-MM'))-error.log"
    }

    "200-033-hisdb-vghtc-ALL-error" = @{
        "root_path" = "\\172.20.1.122\log\server-log\200-033-hisdb-vghtc\"
        "date_path" = "$($today.tostring('yyyy'))\$($today.ToString('yyyy-MM'))\"
        "file_name" = "$($today.ToString('yyyy-MM'))-ALL-error.log"
    }
}



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


function parser-serverlog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$path
    )

    $path = "C:\temp\2024-02-error.log"
    $check_existed = test-path -path $path

    if ($check_existed) {

        $log_content = Get-Content -Path $path

        $today = get-date
        $compare_str = "============================Date: $($today.ToString('yyyy-MM-dd')) ============================"

        
        for ($i = $log_content.count - 1; $i -ge 0; $i--) {
            # 3種狀況:
            # 1.pass: 最後一行就符合
            # 2.fail: 不是最後一行才符合, 表示有error log
            # 3.fail: 找不到符合的, 表示log沒產生.

            if (($log_content[$i] -eq $compare_str) -and ($i -eq ($log_content.count - 1))) {
                # 1.pass
                $result = @{
                    "result"   = "Pass"
                    "errormsg" = "None"
                }
                return $result
            }
            elseif ($log_content[$i] -eq $compare_str) {
                # 2.fail
                #把log中的訊息記下來
                for ($e = $i; $e -lt $log_content.count; $e++) {  
                    $error_msg += $log_content[$e]
                }

                $result = @{
                    "rsult"    = "Fail"
                    "errormsg" = $error_msg
                }
                return $result
            }
            elseif ($i -eq 0) {
                # 3.fail
                $result = @{
                    "result"   = "Fail"
                    "errormsg" = "找不到日期符合的記錄, log可能沒有產生." 
                }
                return $result
            }

        
        }
    }
    else {
        $result = @{
            "result"   = "Fail"
            "errormsg" = "找不到log檔."
        }
    }

}

#parser-log -path C:\temp\2024-02-error.log

$send_msg = "Server log check report`n==" + $today.ToString('yyyyMMdd') + "==`n"
foreach ($Key in $serverlog_checklist.keys) {
    
    $log_path = $serverlog_checklist[$Key]["root_path"] + $serverlog_checklist[$Key]["date_path"] + $serverlog_checklist[$Key]["file_name"]
    Write-Host $log_path

    $result = parser-serverlog -path $log_path
    if ($result['result'] -eq "Pass") {
        $msg = "🟢 Pass: " + $Key + "`n" + 
        "------------ `n"
    } else {
        $msg = "💩 Fail: " + $Key + "`n" +
        "err log: " + $result['errormsg'] + "`n"
        "------------ `n"

    }

    $send_msg += $msg
    
}

Send-LineNotifyMessage -Message $send_msg