#機房檢查
#系統備份檢查



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
        [string]$size,

        [int16]$size_range = 10 # 預設10%
    )

    $today = Get-Date -Format "yyyyMMdd"
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


    $result = @{
        "file_path"        = "none"
        "file_existed"     = "none"
        "file_date"        = "none"
        "file_datechecked" = "none"
        "file_size"        = "none"
        "file_sizechecked" = "none"
    }
    
    switch ($mode) {
        "yyyyMMdd" { 
            $full_path = "$path\$pre_filename$today.$sub_filename"
            $result["file_path"] = $full_path
        }
        
        "ddd" {
            $full_path = "$path\$pre_filename($today_ofweek).$sub_filename"
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
                if ($targetfile["file_size"] -gt 0) {
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
        $result["file_existed"] = "Fail"
    }

    return $result
}


$check_list = @{

    #"root_path" = "\\172.20.1.122\backup\"

    "hisdb"         = @{
        "root_path" = "\\172.20.1.122\backup"
        "folder"       = "001-002-hisdb"
        "pre_filename" = "hahaha"
        "sub_filename" = "zip"
        "size"         = "40GB"
    }

    "dbSTUDY"       = @{
        "root_path" = "\\172.20.1.122\backup"
        "folder"       = "001-014-dbSTUDY"
        "pre_filename" = "hahaha"
        "sub_filename" = "zip"
        "size"         = "40GB"
    }

    "homecare"      = @{
        "root_path" = "\\172.20.1.122\backup"
        "folder"       = "001-016-homecare"
        "pre_filename" = "hahaha"
        "sub_filename" = "zip"
        "size"         = "40GB"
    }

    "nurse"         = @{
        "root_path" = "\\172.20.1.122\backup"
        "folder"       = "001-025-nurse"
        "pre_filename" = "hahaha"
        "sub_filename" = "zip"
        "size"         = "40GB"
    }


    "pts"           = @{
        "root_path" = "\\172.20.1.122\backup"
        "folder"       = "001-025-pts"
        "pre_filename" = "hahaha"
        "sub_filename" = "zip"
        "size"         = "40GB"
    }

    "sk02p"         = @{
        "root_path" = "\\172.20.1.122\backup"
        "folder"       = "001-067-sk02p"
        "pre_filename" = "hahaha"
        "sub_filename" = "zip"
        "size"         = "40GB"
    }

    "hisdb-vghtc_1" = @{
        "root_path" = "\\172.20.1.122\backup"
        "folder"       = "200-033-hisdb-vghtc"
        "pre_filename" = "hahaha"
        "sub_filename" = "zip"
        "size"         = "40GB"
    }

    "hisdb-vghtc_2" = @{
        "root_path" = "\\172.20.1.122\backup"
        "folder"       = "200-033-hisdb-vghtc"
        "pre_filename" = "hahaha"
        "sub_filename" = "zip"
        "size"         = "40GB"
    }

    "hisdb-vghtc_3" = @{
        "root_path" = "\\172.20.1.122\backup"
        "folder"       = "200-033-hisdb-vghtc"
        "pre_filename" = "hahaha"
        "sub_filename" = "zip"
        "size"         = "40GB"
    }

}



foreach ($items in $check_list.Keys) {
Write-Host $check_list[$items]["Folder"]
}

#check_backup_file -path "\\172.20.1.122\backup\001-014-dbSTUDY" -mode xxx -pre_filename hello -sub_filename zip -size 100

