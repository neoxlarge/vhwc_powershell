# �إ�session��Ÿq���ݮୱ�D�� 172.19.1.24
# �b24����I�ϵ{���s�^���a


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


$Username = "vhcy\73058"
$Password = "Q1220416-"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

$remote_computer = 'remote_WIN2016.vhcy.gov.tw'
$output_path ='\\172.20.5.185\mis\webdriver'

$script_block = {
    param($output_path)

    write-output $output_path
    #powershell���ݵn�J��, ���|��\\172.20.5.185\mis���v��, �n���W�Ϻо���~���v��.  
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


