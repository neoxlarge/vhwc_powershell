# 檢查主機系綷LOG備份
# 檢查4個log檔:
# 1.檢查檔案路徑是否正確, 檔名以當日日期產生.
# 2.檢查log內容, 待合固定的格式就pass. 有其他訊息就當fail, 並將訊息傳line.
#
# 檢查NTP-log
# 1.檢查\\172.20.1.122\log\ntp-log\allntp-ntpsync-yyyyMMdd.txt
# 2.檢查IP下有 "校時結束" 字串, 表pass.
#
# 20240608
# 1. 因為NAS帳號和密碼有改, 連增連線NAS的帳號和密碼.
# 2. 更改檔名為dutycheck-serverlog.ps1  

# 連線網路磁碟機 20240608
$Username = "wmis-000-pc05\vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

$psdname = "server_log"
$psdpath = "\\172.20.1.122"

New-PSDrive -Name $psdname -Root "$psdpath" -PSProvider FileSystem -Credential $credential

#檢查主機系綷LOG備份
################################################################################################################################################
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

        # 20240301,當log內容只有一行時, 會出現錯誤,$log_content[0]會傳回單一字元, 而不是整行, 用@()預先轉成陣列可解.
        # 
        $log_content = @(Get-Content -Path $path)

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
    }
    else {
        $msg = "🚨 Fail: " + $Key + "`n" +
        "err log: " + $result['errormsg'] + "`n"
        "------------ `n"

    }

    $send_msg += $msg
    
}

Send-LineNotifyMessage -Message $send_msg

#ntp-log 檢查
##########################################################################################################################################

$ntp_logpath = "\\172.20.1.122\log\ntp-log\allntp-ntpsync-$($today.AddDays(-1).ToString('yyyyMMdd')).txt" #log產生為前一天晚上, 調整-1天.

$ntp = Get-Content -Path $ntp_logpath

$ntp_result = [ordered]@{}
$index = 0
foreach ($n in $ntp) {
    if ($n -like "*-----172*") {
        #每找到一筆ip,都是建立一筆新記錄. "ip"有2種格式, 一種是172.20.1.1 , 一種是172-20-1-1
        $ntp_result.Add( "$index", @{
                'ip' = $n.Replace('"', '').Replace(' ', '').Replace('-----', '').Replace('-', '.')
                'start'                   = $n.ReadCount
                'ntp'                     = ''
                'end'                     = ''
                'result'                  = 'Fail'
            })
        if ($index -gt 0) {    #每一筆ip, 都是上一筆記錄的結束, 除了第一筆和最後一筆. 這裡排除第一筆.                 
            $ntp_result[$($index - 1)]['end'] = $n.ReadCount - 2  #往上移2行為上一筆的end行
        }

        $index += 1

    } elseif ($n.ReadCount -eq $ntp.count) { #這是最後一行了, 就是最後一筆的end
        $ntp_result[$index - 1]['end'] = $n.ReadCount
        
    } elseif ($n -like "*校時結束*") {  #找到校時結束字串為pass
        $ntp_result[$index - 1]['ntp'] = $n.ReadCount
        $ntp_result[$index - 1]['result'] = 'Pass'
    }
    
}

#Line notify 傳送字數有限制, 分成幾筆傳一次, 以$group決定幾筆.
$msg_title = "NTP chect report ($($ntp_result.keys.count)個)`n ==" + $today.ToString('yyyyMMdd') + "==`n"
$msgs = ""
$group = 10
$counter = 1
foreach ($re in $ntp_result.keys) {
    
    if ($ntp_result[$re]['result'] -eq "pass") {
        $msg = "🟢($re) Pass: " + $ntp_result[$re]['ip'] + "`n"
    }
    else {
        $msg = "🚨($re) Fail: " + $ntp_result[$re]['ip'] + "`n"
    }
       
    $msgs += $msg
    
    if ($counter -eq $group -or ($re -eq $ntp_result.keys.count - 1)) {
        
        Send-LineNotifyMessage -Message $($msg_title + $msgs)
        Start-Sleep -s 1
        $counter = 0
        $msgs = ""
    }
    
    $counter += 1

}

# 釋放網路連線
Remove-PSDrive -Name $psdname

