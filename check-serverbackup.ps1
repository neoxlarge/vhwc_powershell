# 系統備份檢查
# 請先連上72.20.1.122
# 檢查項目
# 1. 檢查檔案路徑是不正確, 檔名結尾有3種, 依當日轉換.
# 2. 檢查檔案的最後存取日期是否為當天.
# 3. 檢查檔案大小, 部分檔案似乎有固定大小, 取20240215的大小, 檢查範圍10%以內. 
#
# 200-033-hisdb-vghtc dmp產生時間為早上11點多, 早於這時會檢查到前一天的檔當而出現Fail判斷. 排程下午1點半執行.
# 排程 powershell.exe -file d:\mis\vhwc_powershell\check-serverbackup.ps1



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


function check_backup_file {
    param(
        [Parameter(Mandatory = $true)]
        [string]$mode,

        [Parameter(Mandatory = $true)]
        [string]$path,

        [Parameter(Mandatory = $true)]
        [string]$pre_filename,

        [Parameter(Mandatory = $true)]
        [string]$sub_filename,

        [Parameter(Mandatory = $true)]
        [int64]$size,
        
        [Parameter(Mandatory = $true)]
        [int]$day_shift,

        [int16]$size_range = 10 # 預設10%
    )

    $today = Get-Date -format "yyyyMMdd"
    $today_ofweek = (Get-Date).DayOfWeek.ToString().Substring(0, 3)
    $today_ofweek_chinese = @{
        "Sun" = "星期日"
        "Mon" = "星期一"
        "Tue" = "星期二"
        "Wed" = "星期三"
        "Thu" = "星期四"
        "Fri" = "星期五"
        "Sat" = "星期六"
    }


    $result = [ordered]@{
        "file_path"        = "none"
        "file_existed"     = "none"
        "file_date"        = "none"
        "file_datechecked" = "none"
        "file_size"        = "none"
        "file_sizechecked" = "none"
    }
    
    switch ($mode) {
        # 對應檔名結尾的3種模式.
        "yyyyMMdd" { 
            $today_eof = ((Get-Date).AddDays($day_shift)).ToString("yyyyMMdd") #PACS系統的備份檔名是前一天的日期, 所以加上 $day_shift 修改到前一天.$today_eof(end of filename)會成為檔名的結尾.
            $full_path = "$path\$pre_filename$today_eof.$sub_filename"
            $result["file_path"] = $full_path
        }
        
        "ddd" {
            $full_path = "$path\$pre_filename$today_ofweek.$sub_filename"
            $result["file_path"] = $full_path
        } 

        "XXX" {
            $full_path = "$path\$pre_filename$($today_ofweek_chinese[$today_ofweek]).$sub_filename"
            $result["file_path"] = $full_path
        }
        Default {
            throw "mode參數不正確性!!"

        }
    }
    
    # 檢查檔案是否存在
    if (test-path -Path $result["file_path"]) {
        #檔案存在
        $result["file_existed"] = "Pass"
        $targetfile = get-item -Path $result["file_path"]
        
        $result["file_date"] = $targetfile.LastWriteTime.ToString("yyyyMMdd")
        # 檢查日期是否正確    
        $check_date = $result["file_date"] -eq $today

        if ($check_date) {
            $result["file_datechecked"] = "Pass"
        }
        else {
            $result["file_datechecked"] = "Fail"
        }

        # 檢查檔案大小

        $result["file_size"] = $targetfile.Length
        switch ($size) {
            0 { 
                #檔案大小不固定
                if ($targetfile.Length -gt 0) {
                    $result["file_sizechecked"] = "Pass"
                }
                else {
                    $result["file_sizechecked"] = "Fail"
                }
            }
            Default {
                #檔案大小較固定
                $check_size = $targetfile.Length -gt $size * ((100 - $size_range) / 100) -and $targetfile.Length -le ($size * (100 + $size_range ) / 100) 
                if ($check_size) {
                    $result["file_sizechecked"] = "Pass"
                }
                else {
                    $result["file_sizechecked"] = "Fail"
                }
            }
        }

    }
    else {
        #檔案不存在
        $result["file_existed"] = "Fail"
    }

    return $result
}


$check_list = [ordered]@{

    "hisdb"         = @{
        "path"         = "\\172.20.1.122\backup\001-002-hisdb"
        "mode"         = "yyyyMMdd"
        "pre_filename" = "exp_full_vhgp_"
        "sub_filename" = "dmp"
        "size"         = 40GB
        "day_shift"    = 0
    }

    "dbSTUDY"       = @{
        "path"         = "\\172.20.1.122\backup\001-014-dbSTUDY"
        "mode"         = "ddd"
        "pre_filename" = "dbSTUDY_"
        "sub_filename" = "zip"
        "size"         = 1577kb
        "day_shift"    = 0
    }

    "homecare"      = @{
        "path"         = "\\172.20.1.122\backup\001-016-homecare"
        "mode"         = "ddd"
        "pre_filename" = "LTC_"
        "sub_filename" = "zip"
        "size"         = 16135kb
        "day_shift"    = 0
    }

    "nurse"         = @{
        "path"         = "\\172.20.1.122\backup\001-025-nurse"
        "mode"         = "xxx"
        "pre_filename" = "nurse_"
        "sub_filename" = "7z"
        "size"         = 111850kb
        "day_shift"    = 0
    }


    "pts"           = @{
        "path"         = "\\172.20.1.122\backup\001-025-pts"
        "mode"         = "xxx"
        "pre_filename" = "report_"
        "sub_filename" = "7z"
        "size"         = 3215kb
        "day_shift"    = 0
    }

    "sk02p"         = @{
        "path"         = "\\172.20.1.122\backup\001-067-sk02p"
        "mode"         = "yyyyMMdd"
        "pre_filename" = "SKImagesH-"
        "sub_filename" = "zip"
        "size"         = 0
        "day_shift"    = -1  #因為PACS系統是前一天的, 移動-1天
    }

    "hisdb-vghtc_1" = @{
        "path"         = "\\172.20.1.122\backup\200-033-hisdb-vghtc"
        "mode"         = "yyyyMMdd"
        "pre_filename" = "exp_lob_vghtc_"
        "sub_filename" = "dmp"
        "size"         = 144000000kb
        "day_shift"    = 0
    }

    "hisdb-vghtc_2" = @{
        "path"         = "\\172.20.1.122\backup\200-033-hisdb-vghtc"
        "mode"         = "yyyyMMdd"
        "pre_filename" = "exp_lob_emr_"
        "sub_filename" = "dmp"
        "size"         = 260800000kb
        "day_shift"    = 0
    }

    "hisdb-vghtc_3" = @{
        "path"         = "\\172.20.1.122\backup\200-033-hisdb-vghtc"
        "mode"         = "yyyyMMdd"
        "pre_filename" = "exp_full_hissp1_"
        "sub_filename" = "dmp"
        "size"         = 16400000kb
        "day_shift"    = 0
    }

}


$check_report = [ordered]@{}

$no = 0  # 200-033-hisdb-vghtc 會有3個一樣的, 加個$no識別.

foreach ($item in $check_list.Keys) {
    $result = check_backup_file -mode $check_list[$item]["mode"] `
        -path $check_list[$item]["path"] `
        -pre_filename $check_list[$item]["pre_filename"] `
        -sub_filename $check_list[$item]["sub_filename"] `
        -size $check_list[$item]["size"] `
        -day_shift $check_list[$item]["day_shift"]
    
    $check_report.Add("$($check_list[$item]['path'].Split('\')[-1])_$no" , $result)   # 200-033-hisdb-vghtc 會有3個一樣的, 加個$no識別.
    
    $no += 1
}

$send_msg = "Server backup check `n== $(get-date -format yyyyMMdd) ==`n"

foreach ($r in $check_report.keys) {

    $check_allpass = $check_report[$r]['file_existed'] -eq "Pass" -and $check_report[$r]['file_datechecked'] -eq "Pass" -and $check_report[$r]['file_sizechecked'] -eq "Pass"

    if ($check_allpass) {
        $msg = "🟢 Pass: " + $r.Split('_')[0] + "`n" +
                #"path: " + $check_report[$r]['file_path'] + "`n" +
                #"date: " + $check_report[$r]['file_date'] + "`n" +
                #"size: " + $check_report[$r]['file_size'] + "`n" +
                "------------ `n"

    } else {
        $msg = "🚨 Fail: " + $r.Split('_')[0] + "`n" +
                "path: " + $check_report[$r]['file_path'].Split('\')[-1] + "`n" +
                "date: " + $check_report[$r]['file_date'] + "`n" +
                "size: " + $check_report[$r]['file_size'] + "`n" +
                "------------ `n"
    }

    $send_msg += $msg
}
 
Send-LineNotifyMessage -Message $send_msg



















