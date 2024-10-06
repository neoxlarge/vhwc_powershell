# �ˬdServer�W���{���O�_���


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
        'account' = 'user';
        'password' = 'acervghtc'
    }
}

$server= $server_list.transform1

$Username = ".\$($server.account)"
$Password = "$server.password"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)


$processes = Get-WmiObject -ComputerName $server.ip -Credential $credential -class win32_process 

$processes = Where-Object -InputObject $processes -FilterScript {$_.Name -in $server.processes}

# 1.�ˬd�{���ƶq�O�_���T, �p�G����, ��X�֨��@��
if ($processes.count -ne $server.processes.count) {
    $missingProcesses = Compare-Object -ReferenceObject $server.processes -DifferenceObject $processes.Name -IncludeEqual -ExcludeDifferent 
    Write-Host "Missing processes: $missingProcesses" -ForegroundColor Red
}
