# �ˬd�������C�L��client�ݬO�_���}���ΰ���
# 1. ��ping�ˬdclient�ݬO�_���}��
# 2. ��port 9100�ˬdclient�ݬO�_������

# ������, ���F�קK�{�����а���, �u�঳�@�Ӱ���b���椤.
$mutexName = "Global\dutycheck-idms"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)
if ($mutex.WaitOne(0, $false) -eq $false) { 
    Write-Host "�������C�L�ˬd�q���v�b���椤,����."
    start-sleep -Seconds 5
    exit 
}


# �]�w��
# debug log����T, Continue�|���, SilentContinue���|���.
$DebugPreference = "Continue"
# �]�wLine Notify��Token
$line_token = "ZAxQqfCDIuTL7MzURX1pKTuciEOqnwMqy8lnHNJXEMF"
# �w�ɪ��ɶ�
$timer_hours = @(8..17) #8�I��17�I
$timer_minutes = @(0, 15, 30, 45)



# �������C�L���Ȥ��
$idsm_clients = @{
    "wmis-000-pc05" = @{
        "ip"   = "172.20.5.185"
        "port" = "2788" # or port 9080
    }

    "wnur-b3w-pc04" = @{
        "ip"   = "172.20.2.97"
        "port" = "2788"
    }
}    

function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "HdkeCg1k4nehNa8tEIrJKYrNOeNZMrs89LQTKbf1tbz", # Line Notify �s���v��

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
        $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payload
        
        # �T�����\�ǰe
        Write-Debug "�T���w���\�ǰe�C"
    }
    catch {
        # �o�Ϳ��~�A��X���~�T��
        Write-Error $_.Exception.Message
    }
}


function schedulecheck-idmsclients {

    param (
        $idms_clients
    )

    # �j���ˬd�C�ӫȤ��
    foreach ($clientName in $idms_clients.Keys) {
        $clientInfo = $idsm_clients[$clientName]
        $ipAddress = $clientInfo.ip
        $portNumber = $clientInfo.port

        # �ˬdIP�O�_�iPing�q
        if (Test-Connection -ComputerName $ipAddress -Count 2 -Quiet) {
            Write-Host "$clientName ($ipAddress) is reachable."
            $idsm_clients[$clientName]["reachable"] = $true # �N���G�g�J $idsm_clients

            # �ˬdPort�A�ȬO�_���^��
            $connectionTestResult = Test-NetConnection -ComputerName $ipAddress -Port $portNumber
            if ($connectionTestResult.TcpTestSucceeded) {
                Write-Host "  - Port $portNumber is open and responding."
                $idsm_clients[$clientName]["portResponding"] = $true # �N���G�g�J $idsm_clients
            }
            else {
                Write-Host "  - Port $portNumber is not responding."
                $idsm_clients[$clientName]["portResponding"] = $false # �N���G�g�J $idsm_clients
            }
        }
        else {
            Write-Host "$clientName ($ipAddress) is not reachable."
            $idsm_clients[$clientName]["reachable"] = $false # �N���G�g�J $idsm_clients
            $idsm_clients[$clientName]["portResponding"] = $false # �N���G�g�J $idsm_clients
        }
    }




}





