# �ˬdServer�W���{���O�_���

$DebugPreference = 'Continue'

$server_list = [ordered]@{
    'transform1' = @{
        'title'        = '���������'
        'computername' = 'Blade64-Srv3-wc';
        'ip'           = '172.20.200.41';
        'processes'    = @('hdste02prj.exe',
            'hdste03prj.exe',
            'hdste04prj.exe',
            'hdste05prj.exe',
            'hdste06prj.exe',
            'hdste07prj.exe',
            'hdstq09prj.exe',
            'hdstq10prj.exe');
        'account'      = 'opdvghtc';
        'password'     = 'acervghtc'
    }

    'transform2' = @{
        'title'        = '��������(�p�޵{��)';
        'computername' = 'wser-005-cloudmed';
        'ip'           = '172.20.1.5';
        'processes'    = @('ep.exe',
            '06-�s��his�t��������i�^�ǵ{��APPPRJ .exe',
            'CTMRIUpload.exe',
            'NHI_EII_View.exe');
        'account'      = 'user';
        'password'     = 'tedpc017E'

    }

    'transform3' = @{
        'title'       = '�ǫO�d1.0 �W�ǵ{��';
        'coputername' = 'wmis-111-pc01';
        'ip'          = '172.20.1.4';
        'processes'   = @('NHI_EII_View.exe',
            'IccPrj.exe',
            'PhrB0O0Prj.exe',
            'RegB090Prj.exe',
            'RegB092Prj.exe',
            'RegB093Prj.exe')
        'account'     = 'user';
        'password'    = 'Us2791072'             
    }

    'tranform4'  = @{
        'title'        = '���ݧ妸�U��';
        'computername' = 'wadm-inx-pc02x';
        'ip'           = '172.20.5.147';
        'processes'    = @('IccPrj.exe',
            'HISLogin.exe',
            'HISSystem.exe')
        'account'      = 'user';
        'password'     = 'Us2791072'                
    }

    'tranform5'  = @{
        'title'        = '��E�q��';
        'computername' = 'wadm-in';
        'ip'           = '172.20.200.49'
        'processes'    = @('ERClient.exe',
            'pycharm64.exe',
            'py.exe')
        'account'      = 'Administrator';
        'password'     = 'Acervghtc!'                

    }

    'tranform6'  = @{
        'title'        = '�|?�鵲�{��';
        'computername' = 'unknown';
        'ip'           = '172.20.1.3';
        'processes'    = @('attprj.exe',
            'NisT010.exe')
        'account'      = 'Administrator';
        'password'     = '279!b4E'

    }


    'tranform7'  = @{
        'title'        = 'ĵ���κa��';
        'computername' = 'clonet21';
        'ip'           = '172.20.200.225';
        'processes'    = @('Atcjob.exe',
            'AutoMailReport.exe',
            'PliVacSFTP.exe',
            'DrugAlcoholAddiction.exe',
            'cmd.exe')
        'account'      = 'vgh00';
        'password'     = 'acervghtc'               
    }
}




function Send-LineNotifyMessage {
    [CmdletBinding()]
    param (
        
        [string]$Token = "AVt3SxMcHhatY2fuG2j6HzKGdb5BOTmrfAlEiBolQOO", # Line Notify �s���v��

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

$datatable = New-Object System.Data.DataTable
$datatable.Columns.add('computername', [string]) | Out-Null
$datatable.Columns.add('ip', [string]) | Out-Null
$datatable.Columns.add('resonseDateTime', [Datetime]) | Out-Null
$datatable.Columns.add('processName', [string]) | Out-Null
$datatable.Columns.add('processid', [string]) | Out-Null
$datatable.Columns.add('workingsetsize', [long]) | Out-Null
$datatable.Columns.add('ThreadCount', [int]) | Out-Null
$datatable.Columns.add('HandleCount', [int]) | Out-Null
$datatable.Columns.add('cpuUsage', [int]) | Out-Null



do {
    


    foreach ($server in $server_list.Keys ) {

        #���ˬdserver���s�u
        $ping = Test-Connection -ComputerName $server_list[$server].ip -Count 1 -Quiet
        if ($ping) {
            write-debug "$($server_list[$server].title): $($server_list[$server].ip) �s�u���\"
        }
        else {
            Write-Debug "$($server_list[$server].title): $($server_list[$server].ip) �s�u����"
            # Send-LineNotifyMessage -Message "$($server_list[$server].title): $($server_list[$server].ip) �s�u����" -StickerPackageId 1 -StickerId 100
            continue
        }

        if ($ping) {

            $Username = ".\$($server_list[$server].account)"
            $Password = "$($server_list[$server].password)"
            $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

            # �s�u�컷�ݹq���è��o���椤���{��
            Write-Debug "�s�u�� $($server_list[$server].ip) �è��o���椤���{��..."
            $processes = Get-WmiObject -ComputerName $server_list[$server].ip -Credential $credential -class win32_process 
            $processes = $processes | Where-Object -FilterScript { $_.Name -in $server_list[$server].processes } | Select-Object -Property processid, name, workingsetsize, ThreadCount, HandleCount
            # 1.�ˬd�{���ƶq�O�_���T, �p�G����, ��X�֨��@��
            if ($processes.count -ne $server_list[$server].processes.count) {
                $missingProcesses = Compare-Object -ReferenceObject $server_list[$server].processes -DifferenceObject $processes.Name #-IncludeEqual -ExcludeDifferent 
                Write-Host "Missing processes: $($missingProcesses.inputobject)" -ForegroundColor Red
                Send-LineNotifyMessage -Message "$($server_list[$server].title): $($server_list[$server].ip) �ʤֵ{��: $($missingProcesses.inputobject)" -StickerPackageId 1 -StickerId 100
            }

            # �s�u�컷�ݹq���è��o���w�{���� CPU �ϥβv
            Write-Debug "�s�u�� $($server_list[$server].ip) �è��o���w�{���� CPU �ϥβv..."
            $cpuUsage = Get-WmiObject -Class Win32_PerfFormattedData_PerfProc_Process -ComputerName $server_list[$server].ip -Credential $credential |
            Where-Object { $_.IDProcess -in $processes.processid } 

            foreach ($process in $processes) {
                write-debug "���b�ˬd $($process.name)..."

                # �b$cpuUsage�����ۦP��processid
                $cpuUsage_match = $cpuUsage | Where-Object { $_.IDProcess -eq $process.processid }
        
                # �p�G���, �N�[�Jdatatable
                if ($cpuUsage_match -ne $null) {
                    $datatable.Rows.Add($server_list[$server].computername, $server_list[$server].ip, (Get-Date), $process.name, $process.processid, $process.workingsetsize, $process.ThreadCount, $process.HandleCount, $cpuUsage_match.PercentProcessorTime) | Out-Null
                }
                else {
                    # �p�G�S���, �N�[�Jdatatable, cpuUsage����J'none'
                    $datatable.Rows.Add($server_list[$server].computername, $server_list[$server].ip, (Get-Date), $process.name, $process.processid, $process.workingsetsize, $process.ThreadCount, $process.HandleCount, 'none') | Out-Null
                }
    
                # ��X���w���{��, �åB����resonseDateTime�Ƨ�, ���X�̷s��2�����
                $sortedtable = $datatable.Select( "processName = '$($process.name)'", "resonseDateTime DESC")
    
                # 2.��X�̷s2�����, �ˬd�p�G workingsetsize, threadcount , handlecount �ƭȳ��@��, 
                # ��ܵ{���i�����F
                $last2rows = $sortedtable | Select-Object -First 2
        
                $last2rows | Format-Table 
        
                if (($last2rows.Count -eq 2) -and ($last2rows[0].processName -eq $last2rows[1].processName) -and ($last2rows[0].workingsetsize -eq $last2rows[1].workingsetsize) -and ($last2rows[0].ThreadCount -eq $last2rows[1].ThreadCount) -and ($last2rows[0].HandleCount -eq $last2rows[1].HandleCount)) {
                    Write-Host "Warning: $($last2rows[0].processName) on $($last2rows[0].computername) may be crashed." -ForegroundColor Yellow
                    Send-LineNotifyMessage -Message "$($last2rows[0].computername): $($last2rows[0].processName) �i�����F" -StickerPackageId 1 -StickerId 100
                }
    
            }

        }


    }

    # ��$datatable�L�j�ɥi����L�h�O����, �u�O�d�̷s��1000�����.
    
    if ($datatable.Rows.Count -gt 1000) {
        $datatable.Rows.RemoveAt(0) | Out-Null
        Write-Debug "datatable.Rows.Count: $($datatable.Rows.Count)"

    }

    Write-Debug "datatable: $($datatable.Rows.count)"

    Start-Sleep -s 300
}
while (
    $true<# Condition that stops the loop if it returns false #>
)
