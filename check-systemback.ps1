#�����ˬd
#�t�γƥ��ˬd



function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI",                 # Line Notify �s���v��

        [Parameter(Mandatory = $true)]
        [string]$Message,               # �n�o�e���T�����e

        [string]$StickerPackageId,      # �n�@�ֶǰe���K�ϮM�� ID

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




$check_list = @("\\172.20.1.122\backup\001-002-hisdb",
                "\\172.20.1.122\backup\001-014-dbSTUDY", 
                "\\172.20.1.122\backup\001-016-homecare",
                "\\172.20.1.122\backup\001-025-nurse",
                "\\172.20.1.122\backup\001-025-pts",
                "\\172.20.1.122\backup\001-067-sk02p",
                "\\172.20.1.122\backup\200-033-hisdb-vghtc"
                )

$path = "\\172.20.1.122\backup\001-002-hisdb"

$today = Get-Date -Format "yyyyMMdd"

$dmp_filename = "exp_full_vhgp_$today.dmp"
                
$latest_file = Get-ChildItem -Path $path | Where-Object -FilterScript {$_.Name -match $dmp_filename}

if ($latest_file) {

    if ($latest_file.Length -gt 40GB) {
        $line_msg = "$latest_file �s�b, �ɮפj�p: $($latest_file.length/(1024*1024*1024))GB."
    } else {
        $line_msg = "$dmp_filename �ɮפp��40GB "
    }
    
} else {
    $line_msg = "$dmp_filename ���s�b!!!"
}


Send-LineNotifyMessage -Message $line_msg

# \\172.20.1.122\backup\001-014-dbSTUDY\dbSTUDY_Mon.zip