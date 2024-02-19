# 檢查主機系綷LOG備份
# 檢查4個log檔:
# 1.檢查檔案路徑是否正確, 檔名以當日日期產生.
# 2.檢查log內容, 待合固定的格式就pass. 有其他訊息就當fail, 並將訊息傳line.
#
# 檢查NTP-log
# 1.檢查\\172.20.1.122\log\ntp-log\allntp-ntpsync-yyyyMMdd.txt
# 2.檢查IP下有 "校時結束" 字串, 表pass.



$today = get-date

$serverlog_checklist = [ordered]@{
    "001-002-hisdb-error"           = @{
        "root_path" = "\\172.20.1.122\log\server-log\001-002-hisdb\"
        "date_path" = "$($today.ToString('yyyy-MM'))\"
        "file_name" = "$($today.ToString('yyyy-MM'))-error.log"
    }

    "001-002-hisdb-ALL-error"       = @{
        "root_path" = "\\172.20.1.122\log\server-log\001-002-hisdb\"
        "date_path" = "$($today.ToString('yyyy-MM'))\"
        "file_name" = "$($today.ToString('yyyy-MM'))-ALL-error.log"
    }

    "200-033-hisdb-vghtc-error"     = @{
        "root_path" = "\\172.20.1.122\log\server-log\200-033-hisdb-vghtc\"
        "date_path" = "$($today.ToString('yyyy-MM'))\"
        "file_name" = "$($today.ToString('yyyy-MM'))-error.log"
    }

    "200-033-hisdb-vghtc-ALL-error" = @{
        "root_path" = "\\172.20.1.122\log\server-log\200-033-hisdb-vghtc\"
        "date_path" = "$($today.ToString('yyyy-MM'))\"
        "file_name" = "$($today.ToString('yyyy-MM'))-ALL-error.log"
    }
}



function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz", # Line Notify 存取權杖

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

        return $result
    }

}


$send_msg = "Server log check report`n==" + $today.ToString('yyyyMMdd') + "==`n"
foreach ($Key in $serverlog_checklist.keys) {
    
    $log_path = $serverlog_checklist[$Key]["root_path"] + $serverlog_checklist[$Key]["date_path"] + $serverlog_checklist[$Key]["file_name"]
    
    $result = parser-serverlog -path $log_path

    if ($result['result'] -eq "Pass") {
        $msg = "🟢 Pass: " + $Key + "`n" + 
        "------------ `n"
    } else {
        $msg = "🚨 Fail: " + $Key + "`n" +
        "err log: " + $result['errormsg'] + "`n"
        "------------ `n"

    }

    $send_msg += $msg
    
}

Send-LineNotifyMessage -Message $send_msg

#ntp-log 檢查

$ntp_logpath = "\\172.20.1.122\log\ntp-log\allntp-ntpsync-$($today.AddDays(-1).ToString('yyyyMMdd')).txt" #log產生為前一天晚上, 調整-1天.

$ntp  = Get-Content -Path $ntp_logpath

$ntp_content = [ordered]@{}
$ntp_index = 0
$index = 0

#先把開始找出來
foreach ($line in $ntp) {

    if ($line -like "*-----172.*") {
        
        $ip = $line.Replace(' ',"").Replace('-',"").Replace('"','')
        $ntp_content.Add($index, @{'start' = $ntp_index
                                    'ip' = $ip})
        $index += 1
    }

    $ntp_index += 1
}

#再依start的行數 找出end行數. 下一個start的上一行為end.
foreach ($start in $ntp_content.keys) {

    if ($start -ge $ntp_content.Count -1) {
        $end = $ntp.count
    } else {
        $end = $ntp_content[$start + 1]['start'] -1
    }

    $ntp_content[$start].add('end',$end)
    $ntp_content[$start].add('ntp','')
}

#在start 和 end 間找" 校時結束"字串為pass, 沒有的話為fail.
foreach ($key in $ntp_content.keys) {
    
    for ($i = $ntp_content[$key]['start']; $i -le $ntp_content[$key]['end'] ; $i++){
        
        if ($ntp[$i] -like "*校時結束*"){
            $ntp_content[$key]['ntp'] = $i
            $ntp_content[$key]['result'] = 'pass'
            break
        } else {
            #$ntp_content[$key]['ntp'] = $i
            $ntp_content[$key]['result'] = 'fail'
        }
    }
}    

$send_msg = "NTP chect report`n ==" + $today.ToString('yyyyMMdd') + "==`n"

foreach ($re in $ntp_content.keys) {
    if ($ntp_content[$re]['result'] -eq "pass") {
        $msg = "🟢 Pass: " + $ntp_content[$re]['ip'] + "`n"

    } else {
        $msg = "🚨 Fail: " + $ntp_content[$re]['ip'] + "`n"
    }

    $send_msg = $send_msg + $msg
}

Send-LineNotifyMessage -Message $send_msg