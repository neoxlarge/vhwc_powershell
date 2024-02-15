#�����ˬd
#�t�γƥ��ˬd



function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI", # Line Notify �s���v��

        [Parameter(Mandatory = $true)]
        [string]$Message, # �n�o�e���T�����e

        [string]$StickerPackageId, # �n�@�ֶǰe���K�ϮM�� ID

        [string]$StickerId              # �n�@�ֶǰe���K�� ID
    )

    # Line Notify API �� URI
    $uri = "https://notify-api.line.me/api/notify"

    # �]�w HTTP Header�A�]�t Line Notify �s���v��
    $headers = @{ "Authorization" = "Bearer $Token" }

    # �]�w�n�ǰe���T�����e
    $payload = @{
        "message" = $Message
    }

    # �p�G�n�ǰe�K�ϡA�[�J�K�ϮM�� ID �M�K�� ID
    if ($StickerPackageId -and $StickerId) {
        $payload["stickerPackageId"] = $StickerPackageId
        $payload["stickerId"] = $StickerId
    }

    try {
        # �ϥ� Invoke-RestMethod �ǰe HTTP POST �ШD
        Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payload

        # �T�����\�ǰe
        Write-Output "�T���w���\�ǰe�C"
    }
    catch {
        # �o�Ϳ��~�A��X���~�T��
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

        [int16]$size_range = 10 # �w�]10%
    )

    $today = Get-Date -Format "yyyyMMdd"
    $today_ofweek = (Get-Date).DayOfWeek.ToString().Substring(0, 3)
    $today_ofweek_chinese = @{
        "Sun" = "�P����"
        "Mon" = "�P���@"
        "Tue" = "�P���G"
        "Wed" = "�P���T"
        "Thu" = "�P���|"
        "Fri" = "�P����"
        "Sat" = "�P����"
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
            throw "mode�ѼƤ����T��!!"

        }
    }
    
    # �ˬd�ɮ׬O�_�s�b
    if (test-path -Path $result["file_path"]) {
        
        $result["file_existed"] = "Pass"
        
        $targetfile = get-item -Path $result["file_path"]

        $result["file_date"] = $targetfile.LastWriteTime.ToString("yyyyMMdd")
        
        # �ˬd����O�_���T    
        $check_date = $result["file_date"] -eq $today

        if ($check_date) {
            $result["file_datechecked"] = "Pass"
        }
        else {
            $result["file_datechecked"] = "Fail"
        }

        # �ˬd�ɮפj�p

        $result["file_size"] = $targetfile.Length
        switch ($size) {
            0 { 
                #�ɮפj�p���T�w
                if ($targetfile["file_size"] -gt 0) {
                    $result["file_sizechecked"] = "Pass"
                }
                else {
                    $result["file_sizechecked"] = "Fail"
                }
            }
            Default {
                #�ɮפj�p���T�w
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

