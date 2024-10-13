# 檢查Server上的程式是否當機

$DebugPreference = 'Continue'

$server_list = [ordered]@{
    'check_process1' = @{
        'title'        = '檢驗科轉檔';
        'computername' = 'Blade64-Srv3-wc';
        'ip'           = '172.20.200.41';
        'processes'    = @{
            'hdste01prj.exe' = @{
                'processname' = 'hdste01prj.exe';
                'port'        = $null;
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true;
            }

            'hdste02prj.exe' = @{
                'processname' = 'hdste02prj.exe';
                'port'        = $null;
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true;
            }

            'hdste03prj.exe' = @{
                'processname' = 'hdste03prj.exe';
                'port'        = $null;
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true;
            }

            'hdste04prj.exe' = @{
                'processname' = 'hdste04prj.exe';
                'port'        = $null;
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true;
            }

            'hdste05prj.exe' = @{
                'processname' = 'hdste05prj.exe';
                'port'        = $null;
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true;

            }
            
            'hdste06prj.exe' = @{
                'processname' = 'hdste06prj.exe'
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }

            'hdste07prj.exe' = @{
                'processname' = 'hdste07prj.exe'
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }

            'hdstq09prj.exe' = @{
                'processname' = 'hdstq09prj.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            'hdstq10prj.exe' = @{
                'processname' = 'hdstq10prj.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
        }
            
        'account'      = 'opdvghtc';
        'password'     = 'acervghtc'
    }

    'check_process2' = @{
        'title'        = '儀器轉檔(小豬程式)';
        'computername' = 'wser-005-cloudmed';
        'ip'           = '172.20.1.5';
        'processes'    = @{
            'ep.exe'                        = @{
                'processname' = 'ep.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true;
            }

            '06-新舊his系統檢驗報告回傳程式APPPRJ .exe' = @{
                'processname' = '06-新舊his系統檢驗報告回傳程式APPPRJ .exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }

            'CTMRIUpload.exe'               = @{
                'processname' = 'CTMRIUpload.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
                   
            }
            
            'NHI_EII_View.exe'              = @{
                'processname' = 'NHI_EII_View.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
        };
        'account'      = 'user';
        'password'     = 'tedpc017E'

    }

    'transform3'     = @{
        'title'       = '傳保卡1.0 上傳程式';
        'coputername' = 'wmis-111-pc01';
        'ip'          = '172.20.1.4';
        'processes'   = @{
            'NHI_EII_View.exe' = @{
                'processname' = 'NHI_EII_View.exe'
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            
            'IccPrj.exe'       = @{
                'processname' = 'IccPrj.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            'PhrB0O0Prj.exe'   = @{
                'processname' = 'PhrB0O0Prj.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }

            'RegB090Prj.exe'   = @{
                'processname' = 'RegB090Prj.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            'RegB092Prj.exe'   = @{
                'processname' = 'RegB092Prj.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            'RegB093Prj.exe'   = @{
                'processname' = 'RegB093Prj.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
        }
        'account'     = 'user';
        'password'    = 'Us2791072'             
    }

    'tranform4'      = @{
        'title'        = '雲端批次下載';
        'computername' = 'wadm-inx-pc02x';
        'ip'           = '172.20.5.147';
        'processes'    = @{
            'IccPrj.exe'    = @{
                'processname' = 'IccPrj.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true

            }

            'HISLogin.exe'  = @{
                'processname' = 'HISLogin.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            'HISSystem.exe' = @{
                'processname' = 'HISSystem.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
        }
        'account'      = 'user';
        'password'     = 'Us2791072'                
    }

    'tranform5'      = @{
        'title'        = '急診通報';
        'computername' = 'wadm-in';
        'ip'           = '172.20.200.49'
        'processes'    = @{
            'ERClient.exe'  = @{
                'processname' = 'ERClient.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            'pycharm64.exe' = @{
                'processname' = 'pycharm64.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }

            'py.exe'        = @{
                'processname' = 'py.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
        }
        'account'      = 'Administrator';
        'password'     = 'Acervghtc!'                

    }

    'tranform6'      = @{
        'title'        = '會抛日結程式';
        'computername' = 'unknown';
        'ip'           = '172.20.1.3';
        'processes'    = @{
            'attprj.exe'  = @{
                'processname' = 'attprj.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }

            'NisT010.exe' = @{
                'processname' = 'NisT010.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
        }
        'account'      = 'Administrator';
        'password'     = '279!b4E'

    }


    'tranform7'      = @{
        'title'        = '警消及榮民眷資料下載及回報';
        'computername' = 'clinet21';
        'ip'           = '172.20.200.225';
        'processes'    = @{
            'Atcjob.exe'               = @{
                'processname' = 'Atcjob.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            'AutoMailReport.exe'       = @{
                'processname' = 'AutoMailReport.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            'PliVacSFTP.exe'           = @{
                'processname' = 'PliVacSFTP.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            'DrugAlcoholAddiction.exe' = @{
                'processname' = 'DrugAlcoholAddiction.exe';
                'port'        = $null
                'runInterval' = '900'; #15分鐘
                'runMonitor'  = $true
            }
            #'cmd.exe'
        }
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

# 建位資料表
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
            Send-LineNotifyMessage -Message "🚨 $(get-date) `n$($server_list[$server].title) ($($server_list[$server].ip)) 連線失敗" 
            #continue
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
            if ($processes.count -ne $server_list[$server].processes.keys.count) {
                $missingProcesses = Compare-Object -ReferenceObject $server_list[$server].processes.keys -DifferenceObject $processes.Name #-IncludeEqual -ExcludeDifferent 
                Write-Host "Missing processes: $($missingProcesses.inputobject)" -ForegroundColor Red
                Send-LineNotifyMessage -Message "🚨 $(get-date) `n項目: $($server_list[$server].title) `nip: $($server_list[$server].ip) `n缺少程式: $($missingProcesses.inputobject)" 
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
                    Send-LineNotifyMessage -Message "🚨 $(get-date) `n項目:$($server_list[$server].title) `nip:$($last2rows[0].ip) `n$($last2rows[0].processName) 可能當機了" 
                }
    
            }

        }


    }

    # 當$datatable過大時可能佔過多記憶體, 只保留最新的1000筆資料.
 
    while ($datatable.Rows.count -qt 1000) {
        $datatable.Rows.RemoveAt(0) | Out-Null
        Write-Debug "datatable.Rows.Count: $($datatable.Rows.Count)"
    }


    Start-Sleep -s 900
}
while (
    $true<# Condition that stops the loop if it returns false #>
)
