# 檢查Server上的程式是否當機

$DebugPreference = 'Continue'

$server_list = [ordered]@{
    'transform1' = @{
        'title'        = '檢驗科轉檔'
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
        'title'        = '儀器轉檔(小豬程式)';
        'computername' = 'wser-005-cloudmed';
        'ip'           = '172.20.1.5';
        'processes'    = @('ep.exe',
            '06-新舊his系統檢驗報告回傳程式APPPRJ .exe',
            'CTMRIUpload.exe',
            'NHI_EII_View.exe');
        'account'      = 'user';
        'password'     = 'tedpc017E'

    }

    'transform3' = @{
        'title'       = '傳保卡1.0 上傳程式';
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
        'title'        = '雲端批次下載';
        'computername' = 'wadm-inx-pc02x';
        'ip'           = '172.20.5.147';
        'processes'    = @('IccPrj.exe',
            'HISLogin.exe',
            'HISSystem.exe')
        'account'      = 'user';
        'password'     = 'Us2791072'                
    }

    'tranform5'  = @{
        'title'        = '急診通報';
        'computername' = 'wadm-in';
        'ip'           = '172.20.200.49'
        'processes'    = @('ERClient.exe',
            'pycharm64.exe',
            'py.exe')
        'account'      = 'Administrator';
        'password'     = 'Acervghtc!'                

    }

    'tranform6'  = @{
        'title'        = '會?日結程式';
        'computername' = 'unknown';
        'ip'           = '172.20.1.3';
        'processes'    = @('attprj.exe',
            'NisT010.exe')
        'account'      = 'Administrator';
        'password'     = '279!b4E'

    }


    'tranform7'  = @{
        'title'        = '警消及榮民';
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
        
        [string]$Token = "AVt3SxMcHhatY2fuG2j6HzKGdb5BOTmrfAlEiBolQOO", # Line Notify 存取權杖

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

        #先檢查server的連線
        $ping = Test-Connection -ComputerName $server_list[$server].ip -Count 1 -Quiet
        if ($ping) {
            write-debug "$($server_list[$server].title): $($server_list[$server].ip) 連線成功"
        }
        else {
            Write-Debug "$($server_list[$server].title): $($server_list[$server].ip) 連線失敗"
            # Send-LineNotifyMessage -Message "$($server_list[$server].title): $($server_list[$server].ip) 連線失敗" -StickerPackageId 1 -StickerId 100
            continue
        }

        if ($ping) {

            $Username = ".\$($server_list[$server].account)"
            $Password = "$($server_list[$server].password)"
            $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)

            # 連線到遠端電腦並取得執行中的程式
            Write-Debug "連線到 $($server_list[$server].ip) 並取得執行中的程式..."
            $processes = Get-WmiObject -ComputerName $server_list[$server].ip -Credential $credential -class win32_process 
            $processes = $processes | Where-Object -FilterScript { $_.Name -in $server_list[$server].processes } | Select-Object -Property processid, name, workingsetsize, ThreadCount, HandleCount
            # 1.檢查程式數量是否正確, 如果不對, 找出少那一個
            if ($processes.count -ne $server_list[$server].processes.count) {
                $missingProcesses = Compare-Object -ReferenceObject $server_list[$server].processes -DifferenceObject $processes.Name #-IncludeEqual -ExcludeDifferent 
                Write-Host "Missing processes: $($missingProcesses.inputobject)" -ForegroundColor Red
                Send-LineNotifyMessage -Message "$($server_list[$server].title): $($server_list[$server].ip) 缺少程式: $($missingProcesses.inputobject)" -StickerPackageId 1 -StickerId 100
            }

            # 連線到遠端電腦並取得指定程式的 CPU 使用率
            Write-Debug "連線到 $($server_list[$server].ip) 並取得指定程式的 CPU 使用率..."
            $cpuUsage = Get-WmiObject -Class Win32_PerfFormattedData_PerfProc_Process -ComputerName $server_list[$server].ip -Credential $credential |
            Where-Object { $_.IDProcess -in $processes.processid } 

            foreach ($process in $processes) {
                write-debug "正在檢查 $($process.name)..."

                # 在$cpuUsage中找到相同的processid
                $cpuUsage_match = $cpuUsage | Where-Object { $_.IDProcess -eq $process.processid }
        
                # 如果找到, 就加入datatable
                if ($cpuUsage_match -ne $null) {
                    $datatable.Rows.Add($server_list[$server].computername, $server_list[$server].ip, (Get-Date), $process.name, $process.processid, $process.workingsetsize, $process.ThreadCount, $process.HandleCount, $cpuUsage_match.PercentProcessorTime) | Out-Null
                }
                else {
                    # 如果沒找到, 就加入datatable, cpuUsage欄位填入'none'
                    $datatable.Rows.Add($server_list[$server].computername, $server_list[$server].ip, (Get-Date), $process.name, $process.processid, $process.workingsetsize, $process.ThreadCount, $process.HandleCount, 'none') | Out-Null
                }
    
                # 找出指定的程式, 並且按照resonseDateTime排序, 取出最新的2筆資料
                $sortedtable = $datatable.Select( "processName = '$($process.name)'", "resonseDateTime DESC")
    
                # 2.找出最新2筆資料, 檢查如果 workingsetsize, threadcount , handlecount 數值都一樣, 
                # 表示程式可能當機了
                $last2rows = $sortedtable | Select-Object -First 2
        
                $last2rows | Format-Table 
        
                if (($last2rows.Count -eq 2) -and ($last2rows[0].processName -eq $last2rows[1].processName) -and ($last2rows[0].workingsetsize -eq $last2rows[1].workingsetsize) -and ($last2rows[0].ThreadCount -eq $last2rows[1].ThreadCount) -and ($last2rows[0].HandleCount -eq $last2rows[1].HandleCount)) {
                    Write-Host "Warning: $($last2rows[0].processName) on $($last2rows[0].computername) may be crashed." -ForegroundColor Yellow
                    Send-LineNotifyMessage -Message "$($last2rows[0].computername): $($last2rows[0].processName) 可能當機了" -StickerPackageId 1 -StickerId 100
                }
    
            }

        }


    }

    # 當$datatable過大時可能佔過多記憶體, 只保留最新的1000筆資料.
    
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
