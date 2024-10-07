# 檢查Server上的程式是否當機


$server_list = [ordered]@{
    'transform1' = @{
        'computername' = 'Blade64-Srv3-wc';
        'ip' = '172.20.200.41';
        'processes' = @('hdste02prj.exe',
                        'hdste03prj.exe',
                        'hdste04prj.exe',
                        'hdste05prj.exe',
                        'hdste06prj.exe',
                        'hdste07prj.exe',
                        'hdstq09prj.exe',
                        'hdstq10prj.exe');
        'account' = 'opdvghtc';
        'password' = 'acervghtc'
    }
}

$server= $server_list.transform1

$Username = ".\$($server.account)"
$Password = "$($server.password)"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)


 $datatable = New-Object System.Data.DataTable
 $datatable.Columns.add('computername', [string]) | Out-Null
 $datatable.Columns.add('ip', [string]) | Out-Null
 $datatable.Columns.add('resonseDateTime', [string]) | Out-Null
 $datatable.Columns.add('processName', [string]) | Out-Null
 #$datatable.Columns.add('status', [string]) | Out-Null
 $datatable.Columns.add('processid', [string]) | Out-Null
 $datatable.Columns.add('workingsetsize', [string]) | Out-Null
 $datatable.Columns.add('ThreadCount', [string]) | Out-Null
 $datatable.Columns.add('HandleCount', [string]) | Out-Null
 $datatable.Columns.add('cpuUsage', [string]) | Out-Null

do {
    

$processes = Get-WmiObject -ComputerName $server.ip -Credential $credential -class win32_process 
$processes = $processes |Where-Object -FilterScript {$_.Name -in $server.processes} | Select-Object -Property processid, name, workingsetsize, ThreadCount, HandleCount

# 1.檢查程式數量是否正確, 如果不對, 找出少那一個
if ($processes.count -ne $server.processes.count) {
    $missingProcesses = Compare-Object -ReferenceObject $server.processes -DifferenceObject $processes.Name #-IncludeEqual -ExcludeDifferent 
    Write-Host "Missing processes: $($missingProcesses.inputobject)" -ForegroundColor Red
}

$cpuUsage = Get-WmiObject -Class Win32_PerfFormattedData_PerfProc_Process -ComputerName $server.ip -Credential $credential |
            Where-Object { $_.IDProcess -in $processes.processid } 

foreach ($process in $processes) {
    # 在$cpuUsage中找到相同的processid
    $cpuUsage_match = $cpuUsage | Where-Object { $_.IDProcess -eq $process.processid }
    
    if ($cpuUsage_match -ne $null) {
        $datatable.Rows.Add($server.computername, $server.ip, (Get-Date), $process.name, $process.processid, $process.workingsetsize, $process.ThreadCount, $process.HandleCount, $cpuUsage_match.PercentProcessorTime) | Out-Null
    }
    else {
        $datatable.Rows.Add($server.computername, $server.ip, (Get-Date), $process.name, $process.processid, $process.workingsetsize, $process.ThreadCount, $process.HandleCount, 'none') | Out-Null
    }
}


Clear-Host
$datatable | format-table 
Start-Sleep -s 60
}
while (
    $true<# Condition that stops the loop if it returns false #>
)
