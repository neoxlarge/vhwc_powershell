#line_notify_token  = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI"


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


$line_token = "CclWwNgG6qbD5qx8eO3Oi4ii9azHfolj17SCzIE9UyI"

Send-LineNotifyMessage -Token $line_token -Message "vhwc test line "