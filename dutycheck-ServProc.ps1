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
                        'hdstq10prj.exe',
                        'xxxx_test_xxx.exe');
        'account' = 'opdvghtc';
        'password' = 'acervghtc'
    }
}

$server= $server_list.transform1

$Username = ".\$($server.account)"
$Password = "$($server.password)"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)


$processes = Get-WmiObject -ComputerName $server.ip -Credential $credential -class win32_process 
$processes = $processes |Where-Object -FilterScript {$_.Name -in $server.processes}

# 1.檢查程式數量是否正確, 如果不對, 找出少那一個
if ($processes.count -ne $server.processes.count) {
    $missingProcesses = Compare-Object -ReferenceObject $server.processes -DifferenceObject $processes.Name #-IncludeEqual -ExcludeDifferent 
    Write-Host "Missing processes: $($missingProcesses.inputobject)" -ForegroundColor Red
}
