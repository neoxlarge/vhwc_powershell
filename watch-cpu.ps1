$remoteComputerName = "wnur-opd-pc18"
$processName = "opdnurse"

#$credential = Get-Credential  # 您需要提供遠端電腦的驗證憑證
$Username = "vhcy\vhwcmis"
$Password = "Mis20190610"
$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)



Write-Host "程式名稱: $($process.Name)"

while ($true) {

    $processes = Get-WmiObject -Class Win32_PerfFormattedData_PerfProc_Process -ComputerName $remoteComputerName -Credential $credential |
    Where-Object { $_.Name -eq $processName } |
    Select-Object -Property Name, PercentProcessorTime
    


    if ($processes) {
        foreach ($process in $processes) {
            #Write-Host "程式名稱: $($process.Name) $(get-date).time"
            if ($process.PercentProcessorTime -gt 30) {
                Write-Host "OPDNURSE CPU 使用率: $((get-date).TimeOfDay) $($process.PercentProcessorTime)%" -ForegroundColor Red
                [console]::beep(1000, 3000)
            }
            else {
                Write-Host "OPDNURSE CPU 使用率: $((get-date).TimeOfDay) $($process.PercentProcessorTime)%"
            }
        }
    }
    else {
        Write-Host "找不到名稱為 '$processName' 的程式。"
    }

    Start-Sleep -Seconds 30

}